#include "MonthReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"

#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept>
#include <QProgressDialog>
#include <QtConcurrent>
#include <QVariant>

MonthReportParser::MonthReportParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
{
}

bool MonthReportParser::scanAndParse()
{
    qDebug() << "========== 开始解析月报 ==========";

    // 清空旧数据
    m_queryTasks.clear();
    m_dateFound = false;
    m_baseYearMonth.clear();
    m_baseTime.clear();
    m_currentTime.clear();
    m_dataCache.clear();

    // 查找 #Date1 和 #Date2 标记
    if (!findDateMarker()) {
        QString errMsg = "错误：未找到 #Date1 或 #Date2 标记";
        qWarning() << errMsg;
        emit parseCompleted(false, errMsg);
        return false;
    }

    // 逐行解析
    int totalRows = m_model->rowCount();
    for (int row = 0; row < totalRows; ++row) {
        parseRow(row);
        emit parseProgress(row + 1, totalRows);
    }

    if (m_queryTasks.isEmpty()) {
        QString warnMsg = "警告：未找到任何数据标记";
        qWarning() << warnMsg;
        emit parseCompleted(false, warnMsg);
        return false;
    }


    collectActualDays();

    if (m_actualDays.isEmpty()) {
        qWarning() << "未找到有效日期，跳过预查询";
        emit parseCompleted(true, "解析完成，但未找到有效日期");
        return true;
    }

    setEditState(PREFETCHING);
    m_model->updateEditability();

    // 启动后台预查询
    qDebug() << "========== 开始后台预查询 ==========";
    startAsyncTask(); // 调用基类方法启动后台任务

    QString msg = QString("解析成功：找到 %1 个数据点，数据加载中...")
        .arg(m_queryTasks.size());
    emit parseCompleted(true, msg);

    qDebug() << "========================================";
    return true;
}

bool MonthReportParser::runAsyncTask()
{
    qDebug() << "[后台线程] 日报预查询开始...";
    // analyzeAndPrefetch 包含了取消检查
    return this->analyzeAndPrefetch();
}

bool MonthReportParser::findDateMarker()
{
    int totalRows = m_model->rowCount();
    int totalCols = m_model->columnCount();

    bool foundDate1 = false;
    bool foundDate2 = false;

    for (int row = 0; row < totalRows; ++row) {
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (!cell) continue;

            // ===== 使用 scanText() =====
            QString text = cell->scanText().trimmed();

            // 查找 #Date1:2024-01
            if (isDate1Marker(text)) {
                m_baseYearMonth = extractYearMonth(text);

                if (m_baseYearMonth.isEmpty()) {
                    qWarning() << "年月格式错误:" << text;
                    return false;
                }

                cell->cellType = CellData::DateMarker;
                cell->markerText = text;  // 保存原始标记
                cell->displayValue = text; 

                foundDate1 = true;
            }
            // 查找 #Date2:08:30
            else if (isDate2Marker(text)) {
                m_baseTime = extractTimeOfDay(text);

                if (m_baseTime.isEmpty()) {
                    qWarning() << "时间格式错误:" << text;
                    return false;
                }

                cell->cellType = CellData::TimeMarker;
                cell->markerText = text;                // 保存原始标记
                cell->displayValue = text;

                foundDate2 = true;
            }

            // 如果两个都找到了，可以提前退出
            if (foundDate1 && foundDate2) {
                m_dateFound = true;
                m_baseDate = m_baseYearMonth;
                return true;
            }
        }
    }

    if (!foundDate1) {
        qWarning() << "未找到 #Date1 标记";
    }
    if (!foundDate2) {
        qWarning() << "未找到 #Date2 标记";
    }

    return foundDate1 && foundDate2;
}

QVariant MonthReportParser::formatDisplayValueForMarker(const CellData* cell) const
{
    if (!cell || cell->markerText.isEmpty()) {
        return cell ? cell->displayValue : QVariant();
    }

    // --- 处理 #Date1 (年月) ---
    if (isDate1Marker(cell->markerText)) { // 使用 MonthReportParser 的 isDate1Marker
        QString yearMonth = extractYearMonth(cell->markerText); // 调用 MonthReportParser 的
        QDate date = QDate::fromString(yearMonth + "-01", "yyyy-MM-dd");
        if (date.isValid()) {
            return QString("%1年%2月").arg(date.year()).arg(date.month());
        }
        else {
            qWarning() << "formatDisplayValueForMarker (Month): Invalid YearMonth:" << yearMonth << "from marker" << cell->markerText;
            return cell->markerText;
        }
    }
    // --- 处理 #Date2 (时间) ---
    else if (isDate2Marker(cell->markerText)) { // 使用 MonthReportParser 的 isDate2Marker
        QString timeStr_HHmmss = extractTimeOfDay(cell->markerText); // 调用 MonthReportParser 的
        if (!timeStr_HHmmss.isEmpty()) {
            QTime time = QTime::fromString(timeStr_HHmmss, "HH:mm:ss");
            if (time.isValid()) return time.toString("HH:mm"); // 返回 HH:mm 格式
            else {
                qWarning() << "formatDisplayValueForMarker (Month): Invalid TimeOfDay:" << timeStr_HHmmss << "from marker" << cell->markerText;
                return cell->markerText;
            }
        }
        else {
            qWarning() << "formatDisplayValueForMarker (Month): extractTimeOfDay failed for marker:" << cell->markerText;
            return cell->markerText;
        }
    }
    // --- 处理 #t# (日) ---
    else if (isTimeMarker(cell->markerText)) { // 使用基类的 isTimeMarker
        // 注意：这里不调用 extractTime，而是直接用 extractDay
        int day = extractDay(cell->markerText); // 调用 MonthReportParser 的 extractDay
        if (day > 0) {
            // 验证日期是否有效
            QDate date = QDate::fromString(m_baseYearMonth + QString("-%1").arg(day, 2, 10, QChar('0')), "yyyy-MM-dd");
            if (date.isValid()) {
                return QVariant(day); // 返回数字
            }
            else {
                qWarning() << "formatDisplayValueForMarker (Month): Invalid day for month:" << day << "from marker" << cell->markerText;
                return cell->markerText; // 返回原始标记表示无效
            }
        }
        else {
            qWarning() << "formatDisplayValueForMarker (Month): Invalid day extracted from:" << cell->markerText;
            return cell->markerText; // 返回原始标记表示无效
        }
    }
    // --- 处理 #d# (数据标记，格式化时不改变) ---
    else if (isDataMarker(cell->markerText)) {
        return cell->markerText;
    }

    // 如果以上都不是，返回原始标记
    return cell->markerText;
}

void MonthReportParser::parseRow(int row)
{
    int totalCols = m_model->columnCount();

    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        // ===== 使用 scanText() =====
        QString text = cell->scanText().trimmed();
        if (text.isEmpty()) continue;

        // 遇到 #t# 日期标记（月报中表示"日"）
        if (isTimeMarker(text)) {
            int day = extractDay(text);

            // 验证日期是否有效
            QDate date = QDate::fromString(
                m_baseYearMonth + QString("-%1").arg(day, 2, 10, QChar('0')),
                "yyyy-MM-dd");
            if (!date.isValid()) {
                cell->cellType = CellData::TimeMarker;
                cell->markerText = text;
                cell->displayValue = text; // <-- 修改
                continue;
            }

            // 有效日期，存储日期信息
            m_currentTime = QString("%1").arg(day);

            cell->cellType = CellData::TimeMarker;
            cell->markerText = text;                    // 保存原始标记
            cell->displayValue = text; // <-- 修改
        }
        // 遇到 #d# 数据标记
        else if (isDataMarker(text)) {
            if (m_currentTime.isEmpty()) {
                qWarning() << QString("行%1列%2 缺少日期信息，跳过").arg(row).arg(col);
                continue;
            }

            QString rtuId = extractRtuId(text);
            if (rtuId.isEmpty()) {
                qWarning() << QString("行%1列%2 RTU号为空，跳过").arg(row).arg(col);
                continue;
            }

            cell->cellType = CellData::DataMarker;
            cell->markerText = text;      // 保存原始标记
            cell->rtuId = rtuId;
            cell->displayValue = text;    // 初始显示标记

            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = "";

            m_queryTasks.append(task);
        }
    }
}

QTime MonthReportParser::getTaskTime(const QueryTask& task)
{
    // 月报中，时间信息固定为 m_baseTime
    // 这里只需要返回基准时间即可
    QTime time = QTime::fromString(m_baseTime, "HH:mm:ss");
    if (!time.isValid()) {
        time = QTime::fromString(m_baseTime, "HH:mm");
    }
    return time;
}

QDateTime MonthReportParser::constructDateTime(const QString& date, const QString& time)
{
    // date: "2024-01-15" (年月日)
    // time: "08:30:00" (时分秒)
    QString dateTimeStr = date + " " + time;
    QDateTime result = QDateTime::fromString(dateTimeStr, "yyyy-MM-dd HH:mm:ss");

    if (!result.isValid()) {
        qWarning() << "月报日期时间构造失败:" << dateTimeStr;
    }

    return result;
}

bool MonthReportParser::executeQueries(QProgressDialog* progress)
{

    if (m_queryTasks.isEmpty()) {
        return true;
    }

    qDebug() << "========== 开始填充月报数据 ==========";

    bool cacheReady = analyzeAndPrefetch();

    if (!cacheReady) {
        qWarning() << "查询失败";
        // 不直接返回，继续用现有缓存填充
    }

    if (progress) {
        progress->setRange(0, m_queryTasks.size());
        progress->setLabelText("正在填充月报数据...");
    }

    int successCount = 0;
    int failCount = 0;

    for (int i = 0; i < m_queryTasks.size(); ++i) {
        if (progress && progress->wasCanceled()) {
            for (int j = i; j < m_queryTasks.size(); ++j) {
                m_queryTasks[j].cell->queryExecuted = false;
            }
            return false;
        }

        const QueryTask& task = m_queryTasks[i];

        // 从任务所在行找到日期标记
        int day = 0;
        int totalCols = m_model->columnCount();
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(task.row, col);
            if (cell && cell->cellType == CellData::TimeMarker) {
                QString dayStr = cell->displayValue.toString();  // 使用 displayValue
                day = dayStr.toInt();
                break;
            }
        }

        if (day == 0) {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
            continue;
        }

        // 构造完整日期
        QString fullDate = QString("%1-%2").arg(m_baseYearMonth).arg(day, 2, 10, QChar('0'));

        // 验证日期有效性
        QDate date = QDate::fromString(fullDate, "yyyy-MM-dd");
        if (!date.isValid()) {
            // 无效日期（如2月30日）
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
            continue;
        }

        // 构造时间戳
        QDateTime dateTime = constructDateTime(fullDate, m_baseTime);
        int64_t timestamp = dateTime.toMSecsSinceEpoch();

        float value = 0.0f;
        if (cacheReady && findInCache(task.cell->rtuId, timestamp, value)) {
            task.cell->value = QString::number(value, 'f', 2);
            task.cell->queryExecuted = true;
            task.cell->querySuccess = true;
            successCount++;
        }
        else {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
        }

        if (progress) {
            progress->setValue(i + 1);
        }
        emit queryProgress(i + 1, m_queryTasks.size());
    }

    qDebug() << QString("月报填充完成: 成功 %1, 失败 %2").arg(successCount).arg(failCount);
    emit queryCompleted(successCount, failCount);
    m_model->notifyDataChanged();

    return successCount > 0;
}

QString MonthReportParser::extractTime(const QString& text) const
{
    // 月报中 #t# 后面是日期数字，不是时间
    // 例如：#t#1 表示1日
    QString dayStr = text.mid(3).trimmed();
    return dayStr;
}

// ===== 私有辅助函数 =====

bool MonthReportParser::isDate1Marker(const QString& text) const
{
    return text.startsWith("#Date1:", Qt::CaseInsensitive);
}

bool MonthReportParser::isDate2Marker(const QString& text) const
{
    return text.startsWith("#Date2:", Qt::CaseInsensitive);
}

QString MonthReportParser::extractYearMonth(const QString& text) const
{
    // #Date1:2024-01 → "2024-01"
    QString yearMonth = text.mid(7).trimmed();

    // 验证格式
    QRegularExpression regex("^\\d{4}-\\d{2}$");
    if (!regex.match(yearMonth).hasMatch()) {
        qWarning() << "年月格式错误，期望 yyyy-MM，实际:" << yearMonth;
        return QString();
    }

    // 验证是否为有效月份
    QDate testDate = QDate::fromString(yearMonth + "-01", "yyyy-MM-dd");
    if (!testDate.isValid()) {
        qWarning() << "无效的年月:" << yearMonth;
        return QString();
    }

    return yearMonth;
}

QString MonthReportParser::extractTimeOfDay(const QString& text) const
{
    // #Date2:08:30 → "08:30:00"
    QString timeStr = text.mid(7).trimmed();

    QStringList parts = timeStr.split(":");
    if (parts.size() == 2) {
        // "08:30" → "08:30:00"
        timeStr += ":00";
    }
    else if (parts.size() != 3) {
        qWarning() << "时间格式错误，期望 HH:mm 或 HH:mm:ss，实际:" << timeStr;
        return QString();
    }

    // 验证时间有效性
    QTime time = QTime::fromString(timeStr, "HH:mm:ss");
    if (!time.isValid()) {
        // 尝试宽松格式
        time = QTime::fromString(timeStr, "H:mm:ss");
        if (!time.isValid()) {
            qWarning() << "时间解析失败:" << timeStr;
            return QString();
        }
    }

    return time.toString("HH:mm:ss");
}

int MonthReportParser::extractDay(const QString& text) const
{
    // #t#1 → 1
    // #t#15 → 15
    QString dayStr = text.mid(3).trimmed();

    bool ok;
    int day = dayStr.toInt(&ok);

    if (!ok || day < 1 || day > 31) {
        qWarning() << "日期数字格式错误:" << text;
        return 0;
    }

    return day;
}

// 收集实际日期
void MonthReportParser::collectActualDays()
{
    m_actualDays.clear();

    qDebug() << "========== 开始收集实际日期 ==========";

    for (int row = 0; row < m_model->rowCount(); ++row) {
        for (int col = 0; col < m_model->columnCount(); ++col) {
            CellData* cell = m_model->getCell(row, col);

            // 检查 cellType 而不是 markerText
            if (cell && cell->cellType == CellData::TimeMarker) {
                QString dayMarker = cell->markerText;

                // **安全检查**：如果 markerText 为空，尝试从 displayValue 读取
                if (dayMarker.isEmpty()) {
                    dayMarker = cell->displayValue.toString();
                    qDebug() << QString("  警告：行%1列%2的TimeMarker没有markerText，使用displayValue: %3")
                        .arg(row).arg(col).arg(dayMarker);
                }

                if (dayMarker.isEmpty()) {
                    qWarning() << QString("  跳过空标记: 行%1列%2").arg(row).arg(col);
                    continue;
                }

                int day = extractDay(dayMarker);

                if (day > 0) {
                    QString fullDate = QString("%1-%2")
                        .arg(m_baseYearMonth)
                        .arg(day, 2, 10, QChar('0'));

                    QDate date = QDate::fromString(fullDate, "yyyy-MM-dd");

                    if (date.isValid()) {
                        m_actualDays.insert(day);
                        qDebug() << QString("  收集日期: 行%1列%2, day=%3, fullDate=%4")
                            .arg(row).arg(col).arg(day).arg(fullDate);
                    }
                    else {
                        qWarning() << QString("  无效日期: %1 (行%2列%3)")
                            .arg(fullDate).arg(row).arg(col);
                    }
                }
                else {
                    qWarning() << QString("  无法解析日期数字: %1 (行%2列%3)")
                        .arg(dayMarker).arg(row).arg(col);
                }
            }
        }
    }

    QList<int> sortedDays = m_actualDays.values();
    std::sort(sortedDays.begin(), sortedDays.end());

    qDebug() << QString("收集完成：共 %1 个有效日期").arg(sortedDays.size());

    QString daysStr;
    for (int i = 0; i < sortedDays.size(); ++i) {
        if (i > 0) daysStr += ", ";
        daysStr += QString::number(sortedDays[i]);
        if (i >= 10) {
            daysStr += "...";
            break;
        }
    }
    qDebug() << "日期列表: " << daysStr;
    qDebug() << "==========================================";
}


// 获取最小日期
int MonthReportParser::getMinDay() const
{
    if (m_actualDays.isEmpty()) return 1;
    return *std::min_element(m_actualDays.begin(), m_actualDays.end());
}

// 获取最大日期
int MonthReportParser::getMaxDay() const
{
    if (m_actualDays.isEmpty()) return 1;
    return *std::max_element(m_actualDays.begin(), m_actualDays.end());
}

QList<BaseReportParser::TimeBlock> MonthReportParser::identifyTimeBlocks()
{
    QList<TimeBlock> blocks;

    if (m_actualDays.isEmpty()) {
        qWarning() << "未找到任何有效日期标记";
        return blocks;
    }

    QTime baseTime = QTime::fromString(m_baseTime, "HH:mm:ss");

    // 为每一天创建一个单独的查询块
    QList<int> sortedDays = m_actualDays.values();
    std::sort(sortedDays.begin(), sortedDays.end());

    for (int day : sortedDays) {
        QString dateStr = QString("%1-%2").arg(m_baseYearMonth).arg(day, 2, 10, QChar('0'));

        // 验证日期有效性
        QDate date = QDate::fromString(dateStr, "yyyy-MM-dd");
        if (!date.isValid()) {
            qWarning() << QString("跳过无效日期：%1").arg(dateStr);
            continue;
        }

        TimeBlock block;
        block.startTime = baseTime;                // 08:30:00
        block.endTime = baseTime.addSecs(60);      // 08:31:00（加1分钟）
        block.startDate = dateStr;
        block.endDate = dateStr;

        blocks.append(block);
    }

    return blocks;
}

// 实现日期范围获取
bool MonthReportParser::getDateRange(QString& startDate, QString& endDate)
{
    if (m_currentQueryStartDate.isEmpty() || m_currentQueryEndDate.isEmpty()) {
        return false;
    }

    startDate = m_currentQueryStartDate;
    endDate = m_currentQueryEndDate;
    return true;
}

void MonthReportParser::onRescanCompleted(int newCount, int modifiedCount, int removedCount, 
                                         const QSet<int>& affectedRows)
{
    if (newCount > 0 || modifiedCount > 0 || removedCount > 0) {
        qDebug() << "========== 月报增量更新日期集合 ==========";
        qDebug() << QString("受影响的行数：%1").arg(affectedRows.size());
        
        updateActualDaysIncremental(affectedRows);
        
        // 如果有删除操作，单独验证
        if (removedCount > 0) {
            qDebug() << "  检测到删除操作，验证日期集合完整性...";
            validateActualDays();
        }
        
        qDebug() << QString("更新后的日期总数：%1").arg(m_actualDays.size());
        qDebug() << "==========================================";
    }
}

bool MonthReportParser::analyzeAndPrefetch()
{
    qDebug() << "========== 月报预查询：重新收集日期 ==========";
    collectActualDays();

    // 1. 识别时间块
    QList<TimeBlock> blocks = identifyTimeBlocks();

    // 使用基类中定义的 m_cancelRequested
    if (m_cancelRequested.loadAcquire()) {
        m_lastPrefetchSuccessCount = 0;
        m_lastPrefetchTotalCount = 0;
        return false;
    }

    if (blocks.isEmpty()) {
        qWarning() << "未识别到有效时间块";
        m_lastPrefetchSuccessCount = 0;
        m_lastPrefetchTotalCount = 0;
        return false;
    }

    // 2. 收集所有唯一的RTU
    QSet<QString> uniqueRTUs;
    for (const QueryTask& task : m_queryTasks) {
        uniqueRTUs.insert(task.cell->rtuId);
    }
    QString rtuList = uniqueRTUs.values().join(",");

    // 4. 执行查询
    int successCount = 0;
    int failCount = 0;
    int intervalSeconds = getQueryIntervalSeconds();

    for (int i = 0; i < blocks.size(); ++i) {
        // 使用基类中定义的 m_cancelRequested
        if (m_cancelRequested.loadAcquire()) {
            qDebug() << "后台查询被中断";
            m_lastPrefetchSuccessCount = successCount;
            m_lastPrefetchTotalCount = blocks.size();
            return false;
        }

        const TimeBlock& block = blocks[i];

        qDebug() << QString("执行查询 %1/%2: %3 %4")
            .arg(i + 1)
            .arg(blocks.size())
            .arg(block.startDate)
            .arg(block.startTime.toString("HH:mm"));

        // 使用基类中定义的统一进度信号 taskProgress
        emit taskProgress(i + 1, blocks.size());

        // 在每次查询前，更新当前查询的日期范围
        m_currentQueryStartDate = block.startDate;
        m_currentQueryEndDate = block.endDate;

        bool success = executeSingleQuery(rtuList,
            block.startTime,
            block.endTime,
            intervalSeconds);

        if (success) {
            successCount++;
        }
        else {
            qWarning() << QString("查询失败: %1").arg(block.startDate);
            failCount++;
        }
    }

    m_lastPrefetchSuccessCount = successCount;
    m_lastPrefetchTotalCount = blocks.size();

    // 只要有一次成功，就认为预查询有效
    return successCount > 0;
}

QString MonthReportParser::findTimeForDataMarker(int row, int col)
{
    // ===== 改进：向左查找最近的日期标记 =====
    // 策略：从当前列向左扫描，找到第一个 TimeMarker
    for (int c = col - 1; c >= 0; --c) {
        CellData* cell = m_model->getCell(row, c);
        if (cell && cell->cellType == CellData::TimeMarker) {
            int day = cell->displayValue.toInt();  // 使用 displayValue
            if (day > 0) {
                return QString::number(day);
            }
        }
    }

    // 没找到，记录警告
    qWarning() << QString("数据标记[%1,%2]左侧未找到日期标记").arg(row).arg(col);
    return QString();
}

void MonthReportParser::validateActualDays()
{
    QSet<int> validDays;

    // 快速扫描所有行，收集当前实际存在的日期
    for (int row = 0; row < m_model->rowCount(); ++row) {
        int day = extractDayFromRow(row);
        if (day > 0) {
            QString fullDate = QString("%1-%2")
                .arg(m_baseYearMonth)
                .arg(day, 2, 10, QChar('0'));

            QDate date = QDate::fromString(fullDate, "yyyy-MM-dd");
            if (date.isValid()) {
                validDays.insert(day);
            }
        }
    }

    // 找出需要移除的日期
    QSet<int> daysToRemove = m_actualDays - validDays;

    if (!daysToRemove.isEmpty()) {
        qDebug() << QString("  移除无效日期：%1 个").arg(daysToRemove.size());

        for (int day : daysToRemove) {
            m_actualDays.remove(day);
            qDebug() << QString("    移除日期：%1").arg(day);
        }
    }
}

void MonthReportParser::updateActualDaysIncremental(const QSet<int>& affectedRows)
{
    if (affectedRows.isEmpty()) {
        qDebug() << "  无受影响的行，跳过更新";
        return;
    }

    // 扫描受影响的行，收集新的日期
    QSet<int> newDays;
    for (int row : affectedRows) {
        int day = extractDayFromRow(row);
        if (day > 0) {
            QString fullDate = QString("%1-%2")
                .arg(m_baseYearMonth)
                .arg(day, 2, 10, QChar('0'));

            QDate date = QDate::fromString(fullDate, "yyyy-MM-dd");

            if (date.isValid()) {
                newDays.insert(day);
                if (!m_actualDays.contains(day)) {
                    qDebug() << QString("  新增日期：行%1, day=%2").arg(row).arg(day);
                }
            }
        }
    }

    // 合并到 m_actualDays
    int beforeSize = m_actualDays.size();
    m_actualDays.unite(newDays);
    int afterSize = m_actualDays.size();

    qDebug() << QString("  日期集合更新：%1 -> %2 (新增 %3 个)")
        .arg(beforeSize).arg(afterSize).arg(afterSize - beforeSize);
}

int MonthReportParser::extractDayFromRow(int row) const
{
    for (int col = 0; col < m_model->columnCount(); ++col) {
        const CellData* cell = m_model->getCell(row, col);

        if (cell && cell->cellType == CellData::TimeMarker) {
            QString dayMarker = cell->markerText;

            if (dayMarker.isEmpty()) {
                dayMarker = cell->displayValue.toString();
            }

            if (!dayMarker.isEmpty()) {
                return extractDay(dayMarker);
            }
        }
    }

    return 0;  // 未找到有效日期
}