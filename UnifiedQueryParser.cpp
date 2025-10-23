#include "UnifiedQueryParser.h"
#include "reportdatamodel.h"
#include "TaosDataFetcher.h"
#include <QMessageBox>
#include <QtConcurrent>
#include <QProgressDialog>
#include <cmath>

UnifiedQueryParser::UnifiedQueryParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
{

}

UnifiedQueryParser::~UnifiedQueryParser()
{

}

void UnifiedQueryParser::setTimeRange(const TimeRangeConfig& config)
{
    m_timeConfig = config;
    qDebug() << QString("设置时间范围：%1 ~ %2, 间隔%3秒")
        .arg(config.startTime.toString())
        .arg(config.endTime.toString())
        .arg(config.intervalSeconds);
}

bool UnifiedQueryParser::scanAndParse()
{
    qDebug() << "========== 开始解析统一查询配置 ==========";

    // 从 Model 的 cells 中读取配置（2列：名称+RTU号）
    return loadConfigFromCells();
}

bool UnifiedQueryParser::loadConfigFromCells()
{
    // 清空旧配置
    m_config.columns.clear();

    // 遍历 Model 的前2列，读取配置
    int totalRows = m_model->rowCount();

    for (int row = 0; row < totalRows; ++row) {
        QModelIndex nameIndex = m_model->index(row, 0);
        QModelIndex rtuIndex = m_model->index(row, 1);

        QString displayName = m_model->data(nameIndex, Qt::DisplayRole).toString().trimmed();
        QString rtuId = m_model->data(rtuIndex, Qt::DisplayRole).toString().trimmed();

        // 遇到第一个完全空行就停止
        if (displayName.isEmpty() && rtuId.isEmpty()) {
            break;  // 停止解析
        }

        // 验证完整性
        if (displayName.isEmpty() || rtuId.isEmpty()) {
            qWarning() << QString("第%1行配置不完整").arg(row + 1);
            continue;
        }

        ReportColumnConfig colConfig;
        colConfig.displayName = displayName;
        colConfig.rtuId = rtuId;
        colConfig.sourceRow = row;

        m_config.columns.append(colConfig);
    }

    if (m_config.columns.isEmpty()) {
        qWarning() << "未找到有效的列配置";
        return false;
    }

    qDebug() << QString("配置加载完成：%1 个数据列").arg(m_config.columns.size());
    return true;
}

const QVector<QDateTime>& UnifiedQueryParser::getTimeAxis() const
{
    QMutexLocker locker(&m_dataMutex);
    return m_timeAxis;
}

const QHash<QString, QVector<double>>& UnifiedQueryParser::getAlignedData() const
{
    QMutexLocker locker(&m_dataMutex);
    return m_alignedData;
}

bool UnifiedQueryParser::runAsyncTask()
{
    qDebug() << "========== 统一查询异步任务开始 ==========";

    if (m_cancelRequested.loadAcquire()) return false;

    if (m_config.columns.isEmpty()) {
        qWarning() << " 配置为空，查询终止。";
        return false;
    }

    if (!m_timeConfig.isValid()) {
        qWarning() << " 时间配置无效，查询终止。";
        return false;
    }

    try {
        // ===== 阶段1：生成时间轴 =====
        emit queryStageChanged("正在生成时间轴...");
        QVector<QDateTime> timeAxis = generateTimeAxis();


        if (timeAxis.isEmpty()) {
            qWarning() << " 生成时间轴失败。";
            return false;
        }

        if (m_cancelRequested.loadAcquire()) return false;

        // ===== 阶段2：构造查询地址 =====
        emit queryStageChanged("正在构造查询语句...");
        QString queryAddr = buildQueryAddress();

        if (m_cancelRequested.loadAcquire()) return false;

        // ===== 阶段3：执行数据库查询 =====
        emit queryStageChanged(QString("正在查询数据库(%1 个RTU)...").arg(m_config.columns.size()));

        auto queryStartTime = QDateTime::currentDateTime();

        std::map<int64_t, std::vector<float>> rawData = m_fetcher->fetchDataFromAddress(queryAddr.toStdString());

        auto queryEndTime = QDateTime::currentDateTime();

        if (rawData.empty()) {
            qWarning() << "数据库未返回任何数据。";
        }

        if (m_cancelRequested.loadAcquire()) return false;

        // ===== 阶段4：数据对齐 =====
        emit queryStageChanged(QString("正在对齐数据(%1 个时间点)...").arg(timeAxis.size()));

        auto alignStartTime = QDateTime::currentDateTime();

        // ===== 【关键修复】直接传递 timeAxis 参数 =====
        QHash<QString, QVector<double>> alignedData = alignData(rawData, timeAxis);

        auto alignEndTime = QDateTime::currentDateTime();

        // ===== 阶段5：安全地更新成员变量 =====
        {
            QMutexLocker locker(&m_dataMutex);
            m_timeAxis = timeAxis;
            m_alignedData = alignedData;
        }

        return true;

    }
    catch (const std::exception& e) {
        qWarning() << " 查询执行失败：" << e.what();
        emit databaseError(QString("数据查询失败: %1").arg(e.what()));
        return false;
    }
}

bool UnifiedQueryParser::executeQueries(QProgressDialog* progress)
{
    Q_UNUSED(progress)
        qWarning() << "同步查询已废弃，请使用基类的 startAsyncTask()";
    // 直接返回 false，强制使用异步
    return false;
}

void UnifiedQueryParser::restoreToTemplate()
{
    // 清空时间轴和数据（保留配置）
    m_timeAxis.clear();
    m_alignedData.clear();
    qDebug() << "统一查询数据已清空";
}

QString UnifiedQueryParser::buildQueryAddress()
{
    // 收集所有 RTU 号
    QStringList rtuList;
    for (const auto& col : m_config.columns) {
        rtuList.append(col.rtuId);
    }

    QString rtuPart = rtuList.join(",");
    QString startStr = m_timeConfig.startTime.toString("yyyy-MM-dd HH:mm:ss");
    QString endStr = m_timeConfig.endTime.toString("yyyy-MM-dd HH:mm:ss");

    // 格式：RTU1,RTU2@起始时间~结束时间#间隔秒数
    return QString("%1@%2~%3#%4")
        .arg(rtuPart)
        .arg(startStr)
        .arg(endStr)
        .arg(m_timeConfig.intervalSeconds);
}

QVector<QDateTime> UnifiedQueryParser::generateTimeAxis()
{
    // 【直接复用版本1的 ReportDataModel::generateTimeAxis】
    QVector<QDateTime> result;

    if (!m_timeConfig.isValid()) {
        qWarning() << "时间配置无效";
        return result;
    }

    // 单点查询（间隔为0）
    if (m_timeConfig.intervalSeconds == 0) {
        result.append(m_timeConfig.startTime);
        return result;
    }

    // 时间范围查询
    QDateTime current = m_timeConfig.startTime;
    qint64 totalSeconds = m_timeConfig.startTime.secsTo(m_timeConfig.endTime);
    int estimatedSize = totalSeconds / m_timeConfig.intervalSeconds + 1;
    result.reserve(estimatedSize);

    while (current <= m_timeConfig.endTime) {
        result.append(current);
        current = current.addSecs(m_timeConfig.intervalSeconds);
    }

    return result;
}

QHash<QString, QVector<double>> UnifiedQueryParser::alignData(
    const std::map<int64_t, std::vector<float>>& rawData,
    const QVector<QDateTime>& timeAxis)  // ← 新增参数
{
    qDebug() << "========== 开始数据对齐 ==========";

    QHash<QString, QVector<double>> result;
    QStringList rtuList;
    int matchCount = 0;
    int totalPoints = 0;

    int64_t toleranceMs = 0;

    // 初始化结果集
    for (const auto& col : m_config.columns) {
        rtuList.append(col.rtuId);
        result[col.rtuId] = QVector<double>(timeAxis.size(), std::numeric_limits<double>::quiet_NaN());
        qDebug() << QString("  初始化RTU: %1").arg(col.rtuId);
    }

    // 容错窗口
    if (m_timeConfig.intervalSeconds == 0) {
        toleranceMs = 10000;  // 10秒
    }
    else {
        toleranceMs = (int64_t)(m_timeConfig.intervalSeconds * 1000);
    }
    if (rawData.empty()) {
        qWarning() << "原始数据为空！";
        return result;
    }

    // 检测时间戳单位
    auto firstTs = rawData.begin()->first;
    QDateTime testDt1 = QDateTime::fromSecsSinceEpoch(firstTs);
    QDateTime testDt2 = QDateTime::fromMSecsSinceEpoch(firstTs);


    bool isMilliseconds = (testDt2.date().year() >= 2000 && testDt2.date().year() <= 2100);
    bool isSeconds = (testDt1.date().year() >= 2000 && testDt1.date().year() <= 2100);

    // 3. 遍历时间轴，对齐数据
    for (int i = 0; i < timeAxis.size(); ++i) {  // ← 使用参数
        if (m_cancelRequested.loadAcquire()) {
            break;
        }

        if (i > 0 && i % 100 == 0) {
            emit queryProgressUpdated(i, timeAxis.size());
        }

        int64_t targetTimeMs = timeAxis[i].toMSecsSinceEpoch();
        int64_t targetTimeForCompare = isMilliseconds ? targetTimeMs : (targetTimeMs / 1000);

        // 查找最接近的时间戳
        int64_t closestTime = -1;
        int64_t minDiff = LLONG_MAX;

        for (auto it = rawData.begin(); it != rawData.end(); ++it) {
            int64_t dbTime = it->first;
            int64_t diff = isMilliseconds
                ? qAbs(dbTime - targetTimeForCompare)
                : qAbs((dbTime * 1000) - targetTimeMs);

            if (diff < minDiff) {
                minDiff = diff;
                closestTime = dbTime;
            }
        }

        // 判断是否在容错范围内
        if (closestTime != -1 && minDiff <= toleranceMs) {
            const std::vector<float>& values = rawData.at(closestTime);

            if (i < 3) {
            }

            if (rtuList.size() != (int)values.size()) {
                qWarning() << QString("   RTU数量不匹配！配置=%1, 数据=%2")
                    .arg(rtuList.size()).arg(values.size());
            }

            for (int j = 0; j < rtuList.size() && j < (int)values.size(); ++j) {
                float rawValue = values[j];

                if (std::isnan(rawValue) || std::isinf(rawValue)) {
                    result[rtuList[j]][i] = std::numeric_limits<double>::quiet_NaN();
                    if (i < 3) {
                    }
                }
                else {
                    result[rtuList[j]][i] = static_cast<double>(rawValue);
                    if (i < 3) {
                    }
                }
            }

            matchCount++;
        }
        else {
            if (i < 3) {
                qDebug() << QString("   超出容错范围");
            }
        }

        totalPoints++;
    }

    if (!m_cancelRequested.loadAcquire()) {
        emit queryProgressUpdated(timeAxis.size(), timeAxis.size());
    }

    return result;
}