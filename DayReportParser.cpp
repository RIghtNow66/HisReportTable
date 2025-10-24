#include "DayReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"

#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept>
#include <QProgressDialog>
#include <QtConcurrent>

DayReportParser::DayReportParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
{
}

bool DayReportParser::scanAndParse()
{
    qDebug() << "========== 开始解析日报 ==========";

    // 清空旧数据
    m_queryTasks.clear();
    m_dateFound = false;
    m_baseDate.clear();
    m_currentTime.clear();
    m_dataCache.clear();

    // 查找 #Date 标记
    if (!findDateMarker()) {
        QString errMsg = "错误：未找到 #Date 标记";
        qWarning() << errMsg;
        emit parseCompleted(false, errMsg);
        return false;
    }

    qDebug() << "基准日期:" << m_baseDate;

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

    qDebug() << "解析完成：找到" << m_queryTasks.size() << "个数据点";

    setEditState(PREFETCHING);
    m_model->updateEditability();  // 通知 Model 更新可编辑状态

    // 启动后台预查询
    qDebug() << "========== 开始后台预查询 ==========";
    startAsyncTask();

    QString msg = QString("解析成功：找到 %1 个数据点，数据加载中...")
        .arg(m_queryTasks.size());
    emit parseCompleted(true, msg);

    qDebug() << "========================================";
    return true;
}

bool DayReportParser::runAsyncTask()
{
    qDebug() << "[后台线程] 日报预查询开始...";
    // analyzeAndPrefetch 包含了取消检查
    return this->analyzeAndPrefetch();
}

bool DayReportParser::findDateMarker()
{
    int totalRows = m_model->rowCount();
    int totalCols = m_model->columnCount();

    for (int row = 0; row < totalRows; ++row) {
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (!cell) continue;

            // ===== 使用 scanText() 而不是 value.toString() =====
            QString text = cell->scanText().trimmed();

            if (isDateMarker(text)) {
                m_baseDate = extractDate(text);

                if (m_baseDate.isEmpty()) {
                    return false;
                }

                m_dateFound = true;
                cell->cellType = CellData::DateMarker;
                cell->markerText = text;  // 保存原始标记

                // 设置显示格式
                cell->displayValue = text;

                // 兼容性
                cell->originalMarker = text;

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

    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        // ===== 使用 scanText() =====
        QString text = cell->scanText().trimmed();
        if (text.isEmpty()) continue;

        // 遇到 #t# 时间标记
        if (isTimeMarker(text)) {
            QString timeStr = extractTime(text);
            // 检查提取是否成功
            if(timeStr.isEmpty()) {
                qWarning() << "行" << row << "列" << col << ": 无法从标记提取有效时间:" << text;
                cell->cellType = CellData::TimeMarker; // 仍然标记为 TimeMarker
                cell->markerText = text;
                cell->displayValue = text; // 显示原始错误标记
                cell->originalMarker = text;
                continue; // 跳过 m_currentTime 设置
            }
            m_currentTime = timeStr; // 仍然需要设置 m_currentTime 用于后续 #d#

            cell->cellType = CellData::TimeMarker;
            cell->markerText = text;    // 设置 markerText
            cell->displayValue = text; // <-- 修改：初始 displayValue 等于 markerText
            cell->originalMarker = text; // 兼容性
            continue;
        }
        // 遇到 #d# 数据标记
        else if (isDataMarker(text)) {
            if (m_currentTime.isEmpty()) {
                qWarning() << QString("行%1列%2 缺少时间信息，跳过").arg(row).arg(col);
                continue;
            }

            QString rtuId = extractRtuId(text);
            if (rtuId.isEmpty()) {
                qWarning() << QString("行%1列%2 RTU号为空，跳过").arg(row).arg(col);
                continue;
            }

            cell->cellType = CellData::DataMarker;
            cell->markerText = text;                    // 保存原始标记
            cell->rtuId = rtuId;
            cell->displayValue = text;                  // 初始显示标记
            cell->originalMarker = text;                // 兼容性

            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = "";

            m_queryTasks.append(task);

            qDebug() << QString("  行%1 列%2: RTU=%3, 时间=%4")
                .arg(row).arg(col).arg(rtuId).arg(m_currentTime);
        }
    }
}

QTime DayReportParser::getTaskTime(const QueryTask& task)
{
    int row = task.row;
    int totalCols = m_model->columnCount();

    // 在同一行查找时间标记
    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (cell && cell->cellType == CellData::TimeMarker) {
            QString timeMarker = cell->markerText;
            if (timeMarker.isEmpty()) {
                // 兼容旧数据或意外情况
                timeMarker = cell->displayValue.toString();
            }
            QString timeStr = extractTime(timeMarker);
            QTime time = QTime::fromString(timeStr, "HH:mm:ss");
            if (!time.isValid()) {
                // extractTime 应该总是返回 HH:mm:ss 或空
                // 如果 timeStr 为空或格式仍不对，这里会捕获到
                qWarning() << "getTaskTime: 解析时间失败：" << timeStr;
                // time = QTime::fromString(timeStr, "HH:mm"); // 备用解析（extractTime 已处理）
            }
            return time;
        }
    }

    return QTime();
}

QVariant DayReportParser::formatDisplayValueForMarker(const CellData* cell) const
{
    qDebug() << "[DayParser::format] formatDisplayValueForMarker";
    if (!cell || cell->markerText.isEmpty()) {
        return cell ? cell->displayValue : QVariant(); // 无标记或cell为空，返回当前值
    }

    // --- 处理 #Date ---
    if (isDateMarker(cell->markerText)) { // 使用 DayReportParser 的 isDateMarker
        QString dateStr = extractDate(cell->markerText); // 调用 DayReportParser 的 extractDate
        QDate date = QDate::fromString(dateStr, "yyyy-MM-dd");
        if (date.isValid()) {
            return QString("%1年%2月%3日")
                .arg(date.year())
                .arg(date.month())
                .arg(date.day());
        }
        else {
            qWarning() << "formatDisplayValueForMarker (Day): Invalid date extracted:" << dateStr << "from marker" << cell->markerText;
            return cell->markerText; // 解析失败返回原始标记
        }
    }
    // --- 处理 #t# (时间) ---
    else if (isTimeMarker(cell->markerText)) { // 使用基类的 isTimeMarker
        QString timeStr_HHmmss = extractTime(cell->markerText); // 调用 DayReportParser 的 extractTime
        if (!timeStr_HHmmss.isEmpty()) {
            QTime time = QTime::fromString(timeStr_HHmmss, "HH:mm:ss");
            if (time.isValid()) {
                return time.toString("HH:mm"); // 返回 HH:mm 格式
            }
            else {
                // 这不应该发生，因为 extractTime 内部已经做了校验
                qWarning() << "formatDisplayValueForMarker (Day): Failed to parse time from extracted string:" << timeStr_HHmmss;
                return cell->markerText; // 解析失败返回原始标记
            }
        }
        else {
            // extractTime 返回空，说明原始标记就有问题
            qWarning() << "formatDisplayValueForMarker (Day): extractTime failed for marker:" << cell->markerText;
            return cell->markerText; // 返回原始标记
        }
    }
    // --- 处理 #d# (数据标记，格式化时不改变) ---
    else if (isDataMarker(cell->markerText)) {
        return cell->markerText;
    }

    // 如果以上都不是，返回原始标记
    return cell->markerText;
}

QString DayReportParser::extractTime(const QString& text) const
{
    // 默认实现：#t#0:00 → "00:00:00"
    if (!text.startsWith("#t#", Qt::CaseInsensitive) || text.length() <= 3) {
        qWarning() << "extractTime: Invalid marker text:" << text;
        return QString(); // 返回空表示失败
    }
    // 提取时间：#t#0:00 → "00:00:00"
    QString timeStr = text.mid(3).trimmed();
    QStringList parts = timeStr.split(":");

    if (parts.size() == 2) {
        timeStr += ":00";
    }
    else if (parts.size() != 3) {
        qWarning() << "时间格式错误:" << text;
        return "00:00:00";
    }

    // 尝试多种格式解析
    QTime time = QTime::fromString(timeStr, "H:mm:ss");
    if (!time.isValid()) {
        time = QTime::fromString(timeStr, "HH:mm:ss");
        if (!time.isValid()) {
            qWarning() << "extractTime: Time parsing failed:" << timeStr << "from marker" << text;
            return QString(); // 返回空表示失败
        }
    }

    return time.toString("HH:mm:ss"); // 总是返回标准格式
}

QDateTime DayReportParser::constructDateTime(const QString& date, const QString& time)
{
    QString dateTimeStr = date + " " + time;
    QDateTime result = QDateTime::fromString(dateTimeStr, "yyyy-MM-dd HH:mm:ss");

    if (!result.isValid()) {
        qWarning() << "日期时间构造失败:" << dateTimeStr;
    }

    return result;
}

bool DayReportParser::executeQueries(QProgressDialog* progress)
{

    if (m_queryTasks.isEmpty()) {
        return true;
    }

    bool cacheReady = analyzeAndPrefetch();

    if (!cacheReady) {
        qWarning() << "查询失败";
        // 不直接返回，继续用现有缓存填充
    }

    qDebug() << "========== 开始填充数据 ==========";


    if (progress) {
        progress->setRange(0, m_queryTasks.size());
        progress->setLabelText("正在填充数据...");
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

        QTime time = getTaskTime(task);
        if (!time.isValid()) {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
            continue;
        }

        QDateTime dateTime = QDateTime(
            QDate::fromString(m_baseDate, "yyyy-MM-dd"),
            time
        );
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

    qDebug() << QString("填充完成: 成功 %1, 失败 %2").arg(successCount).arg(failCount);
    emit queryCompleted(successCount, failCount);
    m_model->notifyDataChanged();

    return successCount > 0;
}

bool DayReportParser::isDateMarker(const QString& text) const
{
    return text.startsWith("#Date:", Qt::CaseInsensitive);
}

QString DayReportParser::extractDate(const QString& text) const
{
    QString dateStr = text.mid(6).trimmed();

    QDate date = QDate::fromString(dateStr, "yyyy-MM-dd");
    if (!date.isValid()) {
        qWarning() << "日期格式错误，期望 yyyy-MM-dd，实际:" << dateStr;
        return QString();
    }

    return dateStr;
}

QString DayReportParser::findTimeForDataMarker(int row, int col)
{
    // 在同一行查找时间标记
    int totalCols = m_model->columnCount();
    for (int c = 0; c < totalCols; ++c) {
        CellData* cell = m_model->getCell(row, c);
        if (cell && cell->cellType == CellData::TimeMarker) {
            return cell->displayValue.toString();
        }
    }
    return QString();
}

QList<BaseReportParser::TimeBlock> DayReportParser::identifyTimeBlocks()
{
    qDebug() << "【日报】identifyTimeBlocks() 被调用";
    QList<TimeBlock> blocks;

    if (m_queryTasks.isEmpty()) {
        return blocks;
    }

    // 按时间排序
    QList<QPair<QTime, int>> sortedTasks;

    for (int i = 0; i < m_queryTasks.size(); ++i) {
        QTime time = getTaskTime(m_queryTasks[i]);
        if (time.isValid()) {
            sortedTasks.append(qMakePair(time, i));
        }
    }

    std::sort(sortedTasks.begin(), sortedTasks.end(),
        [](const QPair<QTime, int>& a, const QPair<QTime, int>& b) {
            return a.first < b.first;
        });

    if (sortedTasks.isEmpty()) {
        return blocks;
    }

    // 识别连续块
    TimeBlock currentBlock;
    currentBlock.startTime = sortedTasks[0].first;
    currentBlock.endTime = sortedTasks[0].first;
    currentBlock.taskIndices.append(sortedTasks[0].second);

    const int CONTINUITY_THRESHOLD = 5 * 60;  // 5分钟

    for (int i = 1; i < sortedTasks.size(); ++i) {
        QTime time = sortedTasks[i].first;
        int taskIdx = sortedTasks[i].second;
        int gapSeconds = currentBlock.endTime.secsTo(time);

        if (gapSeconds <= CONTINUITY_THRESHOLD) {
            currentBlock.endTime = time;
            currentBlock.taskIndices.append(taskIdx);
        }
        else {
            blocks.append(currentBlock);
            currentBlock = TimeBlock();
            currentBlock.startTime = time;
            currentBlock.endTime = time;
            currentBlock.taskIndices.append(taskIdx);
        }
    }

    blocks.append(currentBlock);
    return blocks;
}

bool DayReportParser::analyzeAndPrefetch()
{
    // 1. 识别时间块
    QList<TimeBlock> blocks = identifyTimeBlocks();

    if (m_cancelRequested.loadAcquire()) return false;

    if (blocks.isEmpty()) {
        qWarning() << "未识别到有效时间块";
        return false;
    }

    // 2. 收集所有唯一的RTU
    QSet<QString> uniqueRTUs;
    for (const QueryTask& task : m_queryTasks) {
        uniqueRTUs.insert(task.cell->rtuId);
    }
    QString rtuList = uniqueRTUs.values().join(",");

    qDebug() << "RTU数量：" << uniqueRTUs.size();

    // 3. 决策查询策略
    QList<TimeBlock> mergedBlocks;

    if (blocks.size() == 1) {
        mergedBlocks.append(blocks[0]);
    }
    else {
        TimeBlock currentMerged = blocks[0];

        for (int i = 1; i < blocks.size(); ++i) {
            if (shouldMergeBlocks(currentMerged, blocks[i])) {

                currentMerged.endTime = blocks[i].endTime;
                currentMerged.taskIndices.append(blocks[i].taskIndices);
            }
            else {
                mergedBlocks.append(currentMerged);
                currentMerged = blocks[i];
            }
        }
        mergedBlocks.append(currentMerged);
    }

    qDebug() << "查询策略：" << mergedBlocks.size() << "次查询";

    // 4. 执行查询
    int successCount = 0;  // 改用成功计数
    int failCount = 0;     // 失败计数
    int intervalSeconds = getQueryIntervalSeconds();

    for (int i = 0; i < mergedBlocks.size(); ++i) {
        if (m_cancelRequested.loadAcquire()) {
            qDebug() << "后台查询被中断";
            m_lastPrefetchSuccessCount = successCount;
            m_lastPrefetchTotalCount = mergedBlocks.size();
            return false;
        }
        const TimeBlock& block = mergedBlocks[i];

        if (block.isDateRange()) {
            qDebug() << QString("执行查询 %1/%2: %3 %4 ~ %5 %6")
                .arg(i + 1)
                .arg(mergedBlocks.size())
                .arg(block.startDate)
                .arg(block.startTime.toString("HH:mm"))
                .arg(block.endDate)
                .arg(block.endTime.toString("HH:mm"));
        }
        else {
            qDebug() << QString("执行查询 %1/%2: %3 ~ %4")
                .arg(i + 1)
                .arg(mergedBlocks.size())
                .arg(block.startTime.toString("HH:mm"))
                .arg(block.endTime.toString("HH:mm"));
        }


        emit taskProgress(i + 1, mergedBlocks.size());

        bool success = executeSingleQuery(rtuList,
            block.startTime,
            block.endTime,
            intervalSeconds);

        if (success) {
            successCount++;
        }
        else {
            qWarning() << "查询失败";
            failCount++;
        }
    }

    // 记录统计信息
    qDebug() << QString("预查询完成: 成功 %1/%2").arg(successCount).arg(mergedBlocks.size());
    m_lastPrefetchSuccessCount = successCount;
    m_lastPrefetchTotalCount = mergedBlocks.size();

    // 只要有一次成功就返回 true
    return successCount > 0;
}