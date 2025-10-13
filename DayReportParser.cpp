#include "DayReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"
#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept> 
#include <QProgressDialog>

DayReportParser::DayReportParser(ReportDataModel* model, QObject* parent)
    : QObject(parent)
    , m_model(model)
    , m_fetcher(new TaosDataFetcher())
    , m_dateFound(false)
{
}

DayReportParser::~DayReportParser()
{
    delete m_fetcher;
}

// ===== 核心流程 =====

bool DayReportParser::scanAndParse()
{
    qDebug() << "========== 开始解析日报 ==========";

    // 清空旧数据
    m_queryTasks.clear();
    m_dateFound = false;
    m_baseDate.clear();
    m_currentTime.clear();

    // 第1步：查找 #Date 标记
    if (!findDateMarker()) {
        QString errMsg = "错误：未找到 #Date 标记，无法确定查询日期";
        qWarning() << errMsg;
        emit parseCompleted(false, errMsg);
        return false;
    }

    qDebug() << "基准日期:" << m_baseDate;

    // 第2步：逐行解析标记
    int totalRows = m_model->rowCount();

    for (int row = 0; row < totalRows; ++row) {
        parseRow(row);
        emit parseProgress(row + 1, totalRows);
    }

    // 第3步：检查结果
    if (m_queryTasks.isEmpty()) {
        QString warnMsg = "警告：未找到任何 #d# 数据标记";
        qWarning() << warnMsg;
        emit parseCompleted(false, warnMsg);
        return false;
    }

    QString successMsg = QString("解析成功：找到 %1 个待查询单元格").arg(m_queryTasks.size());
    qDebug() << successMsg;
    qDebug() << "========================================";

    emit parseCompleted(true, successMsg);
    return true;
}

bool DayReportParser::findDateMarker()
{
    int totalRows = m_model->rowCount();
    int totalCols = m_model->columnCount();

    for (int row = 0; row < totalRows; ++row) {
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (!cell) continue;

            QString text = cell->value.toString().trimmed();

            if (isDateMarker(text)) {
                m_baseDate = extractDate(text);

                if (m_baseDate.isEmpty()) {
                    // ... (错误处理)
                    return false;
                }

                m_dateFound = true;
                cell->cellType = CellData::DateMarker;
                cell->originalMarker = text;

                // ===== 【核心修正】 开始 =====
                QDate date = QDate::fromString(m_baseDate, "yyyy-MM-dd");
                if (date.isValid()) {
                    // 将显示格式修改为包含年份
                    cell->value = QString("%1年%2月%3日").arg(date.year()).arg(date.month()).arg(date.day());
                }
                // ===== 【核心修正】 结束 =====

                qDebug() << QString("找到 #Date 标记: 行%1 列%2 → %3")
                    .arg(row).arg(col).arg(m_baseDate);

                return true;
            }
        }
    }

    return false;
}

void DayReportParser::parseRow(int row)
{
    int totalCols = m_model->columnCount();

    // 从左到右扫描
    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->value.toString().trimmed();
        if (text.isEmpty()) continue;

        // ===== 情况1：遇到 #t# 时间标记 =====
        if (isTimeMarker(text)) {
            QString timeStr = extractTime(text);
            m_currentTime = timeStr;

            // 标记单元格类型
            cell->cellType = CellData::TimeMarker;
            cell->originalMarker = text;

            // 更新显示值（去掉 #t# 前缀，只显示 HH:mm）
            cell->value = timeStr.left(5);  // "00:00:00" → "00:00"

            qDebug() << QString("  行%1 列%2: 时间标记 %3 → %4")
                .arg(row).arg(col).arg(text).arg(m_currentTime);
        }
        // ===== 情况2：遇到 #d# 数据标记 =====
        else if (isDataMarker(text)) {
            // 检查前置条件
            if (m_currentTime.isEmpty()) {
                qWarning() << QString("警告：行%1列%2 缺少时间信息，跳过").arg(row).arg(col);
                continue;
            }

            QString rtuId = extractRtuId(text);
            if (rtuId.isEmpty()) {
                qWarning() << QString("警告：行%1列%2 RTU号为空，跳过").arg(row).arg(col);
                continue;
            }

            // 标记单元格类型
            cell->cellType = CellData::DataMarker;
            cell->originalMarker = text;
            cell->rtuId = rtuId;

            // 构建查询地址
            QDateTime startTime = constructDateTime(m_baseDate, m_currentTime);
            QString queryPath = buildQueryPath(rtuId, startTime);

            cell->queryPath = queryPath;

            // 添加到查询任务列表
            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = queryPath;
            m_queryTasks.append(task);

            qDebug() << QString("  行%1 列%2: 数据标记 %3 → 查询地址: %4")
                .arg(row).arg(col).arg(text).arg(queryPath);
        }
    }
}

bool DayReportParser::executeQueries(QProgressDialog* progress)
{
    if (m_queryTasks.isEmpty()) {
        qWarning() << "没有待查询的任务";
        return true; // 没有任务也算成功完成
    }

    qDebug() << "========== 开始执行查询 ==========";
    int successCount = 0;
    int failCount = 0;
    int total = m_queryTasks.size();

    for (int i = 0; i < total; ++i) {
        // 【核心修正】在每次循环开始时检查是否已取消
        if (progress && progress->wasCanceled()) {
            qDebug() << "查询被用户取消。";
            return false; // 返回 false 表示被取消
        }

        const QueryTask& task = m_queryTasks[i];

        // 更新进度条文本
        if (progress) {
            progress->setLabelText(QString("正在查询: %1/%2").arg(i + 1).arg(total));
        }

        try {
            QVariant result = querySinglePoint(task.queryPath);
            task.cell->value = result;
            task.cell->queryExecuted = true;
            task.cell->querySuccess = true;
            successCount++;
        }
        catch (const std::exception& e) {
            qWarning() << QString("失败: 行%1列%2 - %3").arg(task.row).arg(task.col).arg(e.what());
            task.cell->value = "ERROR";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
        }

        emit queryProgress(i + 1, total);
    }

    qDebug() << "========================================";
    qDebug() << QString("查询完成: 成功 %1, 失败 %2").arg(successCount).arg(failCount);

    emit queryCompleted(successCount, failCount);

    // 通知模型刷新显示
    m_model->notifyDataChanged();

    return true; // 返回 true 表示正常完成
}

void DayReportParser::restoreToTemplate()
{
    qDebug() << "恢复到模板初始状态...";
    for (const auto& task : m_queryTasks) {
        if (task.cell) {
            // 将单元格的值恢复为其原始标记文本
            task.cell->value = task.cell->originalMarker;
            task.cell->queryExecuted = false;
            task.cell->querySuccess = false;
        }
    }
}


QVariant DayReportParser::querySinglePoint(const QString& queryPath)
{
    // 调用 TaosDataFetcher 执行查询
    auto dataMap = m_fetcher->fetchDataFromAddress(queryPath.toStdString());

    if (dataMap.empty()) {
        throw std::runtime_error("查询无数据");
    }

    // 取第一个时间点的第一个值
    auto firstEntry = dataMap.begin();
    if (firstEntry->second.empty()) {
        throw std::runtime_error("数据为空");
    }

    float value = firstEntry->second[0];

    // 转换为double并保留2位小数
    double result = static_cast<double>(value);
    return QString::number(result, 'f', 2);
}

// ===== 标记识别 =====

bool DayReportParser::isDateMarker(const QString& text) const
{
    return text.startsWith("#Date:", Qt::CaseInsensitive);
}

bool DayReportParser::isTimeMarker(const QString& text) const
{
    return text.startsWith("#t#", Qt::CaseInsensitive);
}

bool DayReportParser::isDataMarker(const QString& text) const
{
    return text.startsWith("#d#", Qt::CaseInsensitive);
}

// ===== 信息提取 =====

QString DayReportParser::extractDate(const QString& text)
{
    // "#Date:2024-01-01" → "2024-01-01"
    QString dateStr = text.mid(6).trimmed();  // 跳过 "#Date:"

    // 验证日期格式
    QDate date = QDate::fromString(dateStr, "yyyy-MM-dd");
    if (!date.isValid()) {
        qWarning() << "日期格式错误，期望 yyyy-MM-dd，实际:" << dateStr;
        return QString();
    }

    return dateStr;
}

QString DayReportParser::extractTime(const QString& text)
{
    // "#t#0:00" → "00:00:00"
    // "#t#13:30" → "13:30:00"

    QString timeStr = text.mid(3).trimmed();  // 跳过 "#t#"

    // 处理不同格式
    QStringList parts = timeStr.split(":");

    if (parts.size() == 2) {
        // "0:00" → "00:00:00"
        timeStr += ":00";
    }
    else if (parts.size() != 3) {
        qWarning() << "时间格式错误:" << text;
        return "00:00:00";
    }

    // 尝试解析并格式化
    QTime time = QTime::fromString(timeStr, "H:mm:ss");
    if (!time.isValid()) {
        qWarning() << "时间解析失败:" << timeStr;
        return "00:00:00";
    }

    // 返回标准格式 "HH:mm:ss"
    return time.toString("HH:mm:ss");
}

QString DayReportParser::extractRtuId(const QString& text)
{
    // "#d#AIRTU034700019" → "AIRTU034700019"
    return text.mid(3).trimmed();
}

// ===== 查询地址构建 =====

QString DayReportParser::buildQueryPath(const QString& rtuId, const QDateTime& baseTime)
{
    // 起始时间：用户指定的时间
    QString startTime = baseTime.toString("yyyy-MM-dd HH:mm:ss");

    // 结束时间：起始时间 + 1分钟
    QDateTime endTime = baseTime.addSecs(60);
    QString endTimeStr = endTime.toString("yyyy-MM-dd HH:mm:ss");

    // 查询格式：RTU号@起始时间~结束时间#间隔(秒)
    // 示例：AIRTU034700019@2024-01-01 00:00:00~2024-01-01 00:01:00#1
    return QString("%1@%2~%3#1")
        .arg(rtuId)
        .arg(startTime)
        .arg(endTimeStr);
}

QDateTime DayReportParser::constructDateTime(const QString& date, const QString& time)
{
    // "2024-01-01" + "00:00:00" → QDateTime
    QString dateTimeStr = date + " " + time;
    QDateTime result = QDateTime::fromString(dateTimeStr, "yyyy-MM-dd HH:mm:ss");

    if (!result.isValid()) {
        qWarning() << "日期时间构造失败:" << dateTimeStr;
    }

    return result;
}