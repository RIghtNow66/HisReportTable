#include "UnifiedQueryParser.h"
#include "reportdatamodel.h"
#include "TaosDataFetcher.h"
#include <QMessageBox>
#include <QProgressDialog>
#include <cmath>

UnifiedQueryParser::UnifiedQueryParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
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

bool UnifiedQueryParser::executeQueries(QProgressDialog* progress)
{
    if (m_config.columns.isEmpty()) {
        QMessageBox::warning(nullptr, "错误", "配置未加载");
        return false;
    }

    if (!m_timeConfig.isValid()) {
        QMessageBox::warning(nullptr, "错误", "时间范围未设置");
        return false;
    }

    // 1. 生成时间轴
    m_timeAxis = generateTimeAxis();

    if (progress) {
        progress->setLabelText("正在查询数据...");
        progress->setRange(0, 100);
        progress->setValue(10);
    }

    // 2. 构造查询地址
    QString queryAddr = buildQueryAddress();
    qDebug() << "查询地址：" << queryAddr;

    if (progress) progress->setValue(20);

    // 3. 执行查询
    std::map<int64_t, std::vector<float>> rawData;
    try {
        rawData = m_fetcher->fetchDataFromAddress(queryAddr.toStdString());
    }
    catch (const std::exception& e) {
        QMessageBox::critical(nullptr, "查询失败", QString("数据库错误：%1").arg(e.what()));
        return false;
    }

    if (progress) progress->setValue(60);

    if (rawData.empty()) {
        QMessageBox::warning(nullptr, "无数据", "查询未返回任何数据");
        return false;
    }

    // 4. 数据对齐
    m_alignedData = alignData(rawData);

    if (progress) progress->setValue(100);

    qDebug() << "查询完成";
    return true;
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
        qDebug() << "生成单点查询时间：" << m_timeConfig.startTime.toString("yyyy-MM-dd HH:mm:ss");
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

    qDebug() << "生成时间轴：" << result.size() << "个时间点";
    return result;
}

QHash<QString, QVector<double>> UnifiedQueryParser::alignData(
    const std::map<int64_t, std::vector<float>>& rawData)
{
    // 【直接复用版本1的 ReportDataModel::alignDataWithInterpolation】
    // 注意：这里简化了参数，因为 RTU 顺序已经在 m_config.columns 中确定

    QHash<QString, QVector<double>> result;
    QStringList rtuList;

    // 初始化结果容器
    for (const auto& col : m_config.columns) {
        rtuList.append(col.rtuId);
        result[col.rtuId] = QVector<double>(m_timeAxis.size(), std::numeric_limits<double>::quiet_NaN());
    }

    // 容错匹配的时间窗口（间隔的一半）
    int64_t tolerance = (int64_t)(m_timeConfig.intervalSeconds * 1000) / 2;
    int matchCount = 0;

    // 对齐数据
    for (int i = 0; i < m_timeAxis.size(); ++i) {
        int64_t targetTime = m_timeAxis[i].toMSecsSinceEpoch();

        // 查找最接近的时间戳
        int64_t closestTime = -1;
        int64_t minDiff = LLONG_MAX;

        for (auto it = rawData.begin(); it != rawData.end(); ++it) {
            int64_t diff = qAbs(it->first - targetTime);
            if (diff < minDiff) {
                minDiff = diff;
                closestTime = it->first;
            }
        }

        // 如果在容错范围内，使用该数据
        if (closestTime != -1 && minDiff <= tolerance) {
            const std::vector<float>& values = rawData.at(closestTime);

            for (int j = 0; j < rtuList.size() && j < (int)values.size(); ++j) {
                result[rtuList[j]][i] = values[j];
            }
            matchCount++;
        }
    }

    qDebug() << QString("数据对齐完成：匹配 %1/%2").arg(matchCount).arg(m_timeAxis.size());
    return result;
}