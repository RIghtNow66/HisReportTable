#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept> 
#include <QProgressDialog>
#include <QtConcurrent> 

#include "DayReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"


DayReportParser::DayReportParser(ReportDataModel* model, QObject* parent)
    : QObject(parent)
    , m_model(model)
    , m_fetcher(new TaosDataFetcher())
    , m_dateFound(false)
    , m_prefetchWatcher(nullptr)  
    , m_isPrefetching(false)     
    , m_stopRequested(0)
{
    //  创建 FutureWatcher
    m_prefetchWatcher = new QFutureWatcher<bool>(this);

    //  连接完成信号
    connect(m_prefetchWatcher, &QFutureWatcher<bool>::finished,
        this, &DayReportParser::onPrefetchFinished);
}

DayReportParser::~DayReportParser()
{
    if (m_prefetchWatcher) {
        disconnect(m_prefetchWatcher, nullptr, this, nullptr);
    }

    if (m_isPrefetching) {
        qDebug() << "请求停止预查询...";
        m_stopRequested.storeRelease(1);  // 设置停止标志

        // 不再阻塞等待，而是异步取消
        if (!m_prefetchFuture.isFinished()) {
            // 尝试取消（注意：QtConcurrent::run不支持取消，只能等待）
            qDebug() << "等待后台线程退出（最多3秒）...";

            // 使用事件循环等待，避免阻塞UI
            QEventLoop loop;
            QTimer::singleShot(3000, &loop, &QEventLoop::quit);

            // 不再连接finished信号，因为已经断开了
            QTimer checkTimer;
            checkTimer.setInterval(100);
            connect(&checkTimer, &QTimer::timeout, [&]() {
                if (m_prefetchFuture.isFinished()) {
                    loop.quit();
                }
                });
            checkTimer.start();

            loop.exec();

            if (!m_prefetchFuture.isFinished()) {
                qWarning() << "后台线程未能及时退出，强制析构";
            }
        }
    }
    delete m_fetcher;
}

void DayReportParser::onPrefetchFinished()
{
    m_isPrefetching = false;

    bool success = m_prefetchFuture.result();

    if (success) {
        QMutexLocker locker(&m_cacheMutex);
        qDebug() << " 预查询完成：已缓存" << m_dataCache.size() << "个数据点";
    }
    else {
        qWarning() << " 预查询失败";
    }
}

void DayReportParser::clearCache()
{
    QMutexLocker locker(&m_cacheMutex);  // 自动加锁
    qDebug() << "清空缓存：" << m_dataCache.size() << "个数据点";
    m_dataCache.clear();
}

// 从任务获取时间
QTime DayReportParser::getTaskTime(const QueryTask& task)
{
    int row = task.row;
    int totalCols = m_model->columnCount();

    // 在同一行查找时间标记
    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (cell && cell->cellType == CellData::TimeMarker) {
            QString timeStr = cell->value.toString();
            QTime time = QTime::fromString(timeStr, "HH:mm:ss");
            if (!time.isValid()) {
                time = QTime::fromString(timeStr, "HH:mm");
            }
            return time;
        }
    }

    return QTime();
}



// ===== 核心流程 =====
bool DayReportParser::findInCache(const QString& rtuId, int64_t timestamp, float& value)
{
    QMutexLocker locker(&m_cacheMutex);

    CacheKey key;
    key.rtuId = rtuId;
    key.timestamp = timestamp;

    // 1. 精确匹配
    if (m_dataCache.contains(key)) {
        value = m_dataCache[key];
        return true;
    }

    // 2. 容错匹配：±60秒内的最近时间点
    const int64_t tolerance = 60000;  // 60秒（毫秒）

    int64_t closestDiff = std::numeric_limits<int64_t>::max();
    bool found = false;  // 是否找到匹配
    float closestValue = 0.0f;  // 最接近的值

    for (auto it = m_dataCache.constBegin(); it != m_dataCache.constEnd(); ++it) {
        if (it.key().rtuId == rtuId) {
            int64_t diff = qAbs(it.key().timestamp - timestamp);
            if (diff <= tolerance) {
                if (diff < closestDiff) {
                    closestDiff = diff;
                    closestValue = it.value();
                    found = true;
                }
            }
        }
    }

    if (found) {
        value = closestValue;
        return true;
    }

    return false;
}

bool DayReportParser::executeSingleQuery(const QString& rtuList,
    const QTime& startTime,
    const QTime& endTime,
    int intervalSeconds)
{
    // 构造查询地址
    QString query = QString("%1@%2 %3~%2 %4#%5")
        .arg(rtuList)
        .arg(m_baseDate)
        .arg(startTime.toString("HH:mm:ss"))
        .arg(endTime.addSecs(60).toString("HH:mm:ss"))  // 结束时间+1分钟
        .arg(intervalSeconds);

    qDebug() << "  查询地址：" << query;

    try {
        // 执行查询
        auto dataMap = m_fetcher->fetchDataFromAddress(query.toStdString());

        qDebug() << "  返回时间点数量：" << dataMap.size();

        if (dataMap.empty()) {
            qWarning() << "  查询无数据";
            return false;
        }
        {
            QMutexLocker locker(&m_cacheMutex);
            // 解析RTU列表
            QStringList rtuArray = rtuList.split(",");

            // 存入缓存
            for (auto it = dataMap.begin(); it != dataMap.end(); ++it) {
                int64_t timestamp = it->first;
                const std::vector<float>& values = it->second;

                for (int i = 0; i < rtuArray.size() && i < values.size(); ++i) {
                    CacheKey key;
                    key.rtuId = rtuArray[i];
                    key.timestamp = timestamp;

                    m_dataCache[key] = values[i];
                }
            }

            qDebug() << QString(" 已缓存 %1 个时间点 × %2 个RTU = %3 个值")
                .arg(dataMap.size())
                .arg(rtuArray.size())
                .arg(m_dataCache.size());
        }
        return true;
    }
    catch (const std::exception& e) {
        qWarning() << "  查询失败：" << e.what();
        return false;
    }
}

bool DayReportParser::shouldMergeBlocks(const TimeBlock& block1, const TimeBlock& block2)
{
    // 计算两个块之间的间隔
    int gapMinutes = block1.endTime.secsTo(block2.startTime) / 60;

    // 经验阈值：间隔超过4小时（240分钟），就不合并
    const int MERGE_THRESHOLD_MINUTES = 2 * 60;

    bool shouldMerge = gapMinutes < MERGE_THRESHOLD_MINUTES;

    if (shouldMerge) {
        qDebug() << QString("  → 间隔%1分钟 < %2分钟，合并")
            .arg(gapMinutes).arg(MERGE_THRESHOLD_MINUTES);
    }
    else {
        qDebug() << QString("  → 间隔%1分钟 >= %2分钟，不合并")
            .arg(gapMinutes).arg(MERGE_THRESHOLD_MINUTES);
    }

    return shouldMerge;
}

QList<DayReportParser::TimeBlock> DayReportParser::identifyTimeBlocks()
{
    QList<TimeBlock> blocks;

    if (m_queryTasks.isEmpty()) {
        return blocks;
    }

    // 按时间排序
    QList<QPair<QTime, int>> sortedTasks;

    for (int i = 0; i < m_queryTasks.size(); ++i) {
        QTime time = getTaskTime(m_queryTasks[i]);  //  使用新方法
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

    if (m_stopRequested.loadAcquire()) return false;

    if (blocks.isEmpty()) {
        qWarning() << "未识别到有效时间块";
        return false;
    }

    qDebug() << "识别到" << blocks.size() << "个时间块：";
    for (int i = 0; i < blocks.size(); ++i) {
        qDebug() << QString("  块%1: %2 ~ %3 (%4个数据点)")
            .arg(i + 1)
            .arg(blocks[i].startTime.toString("HH:mm"))
            .arg(blocks[i].endTime.toString("HH:mm"))
            .arg(blocks[i].taskIndices.size());
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
        // 只有一个块，直接查询
        mergedBlocks.append(blocks[0]);
    }
    else {
        // 多个块，判断是否合并
        TimeBlock currentMerged = blocks[0];

        for (int i = 1; i < blocks.size(); ++i) {
            if (shouldMergeBlocks(currentMerged, blocks[i])) {
                // 合并
                qDebug() << QString("  合并块 %1~%2 和 %3~%4")
                    .arg(currentMerged.startTime.toString("HH:mm"))
                    .arg(currentMerged.endTime.toString("HH:mm"))
                    .arg(blocks[i].startTime.toString("HH:mm"))
                    .arg(blocks[i].endTime.toString("HH:mm"));

                currentMerged.endTime = blocks[i].endTime;
                currentMerged.taskIndices.append(blocks[i].taskIndices);
            }
            else {
                // 不合并，保存当前块，开始新块
                mergedBlocks.append(currentMerged);
                currentMerged = blocks[i];
            }
        }
        mergedBlocks.append(currentMerged);  // 最后一个块
    }

    qDebug() << "查询策略：" << mergedBlocks.size() << "次查询";

    // 4. 执行查询
    bool allSuccess = true;
    for (int i = 0; i < mergedBlocks.size(); ++i) {
        if (m_stopRequested.loadAcquire()) {  // 每次循环检查
            qDebug() << "后台查询被中断";
            return false;
        }
        const TimeBlock& block = mergedBlocks[i];

        qDebug() << QString("执行查询 %1/%2: %3 ~ %4")
            .arg(i + 1)
            .arg(mergedBlocks.size())
            .arg(block.startTime.toString("HH:mm"))
            .arg(block.endTime.toString("HH:mm"));

        emit prefetchProgress(i + 1, mergedBlocks.size());

        bool success = executeSingleQuery(rtuList,
            block.startTime,
            block.endTime,
            60);  // 间隔60秒（因为Time最小单位是分钟）

        if (!success) {
            qWarning() << "查询失败";
            allSuccess = false;
        }
    }

    return allSuccess;
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

    // =====  核心修改：启动后台预查询 =====
    qDebug() << "========== 开始后台预查询 ==========";

    m_isPrefetching = true;

    // 在后台线程执行
    m_prefetchFuture = QtConcurrent::run([this]() -> bool {
        qDebug() << "[后台线程] 预查询开始...";
        try {
            bool result = this->analyzeAndPrefetch();
            qDebug() << "[后台线程] 预查询" << (result ? "成功" : "失败");
            return result;
        }
        catch (const std::exception& e) {
            qWarning() << "[后台线程] 异常：" << e.what();
            return false;
        }
        });

    m_prefetchWatcher->setFuture(m_prefetchFuture);

    //  立即返回，不阻塞UI
    QString msg = QString("解析成功：找到 %1 个数据点，数据加载中...")
        .arg(m_queryTasks.size());
    emit parseCompleted(true, msg);

    qDebug() << "========================================";
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

    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->value.toString().trimmed();
        if (text.isEmpty()) continue;

        // ===== 情况1：遇到 #t# 时间标记 =====
        if (isTimeMarker(text)) {
            QString timeStr = extractTime(text);
            m_currentTime = timeStr;

            cell->cellType = CellData::TimeMarker;
            cell->originalMarker = text;
            cell->value = timeStr.left(5);  // "00:00:00" → "00:00"

            qDebug() << QString("行%1 列%2: 时间标记 %3 → %4")
                .arg(row).arg(col).arg(text).arg(m_currentTime);
        }
        // ===== 情况2：遇到 #d# 数据标记 =====
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

            // 标记单元格
            cell->cellType = CellData::DataMarker;
            cell->originalMarker = text;
            cell->rtuId = rtuId;

            // 添加到任务列表（不再构建复杂的queryPath）
            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = "";  // 新方案不需要单独的queryPath

            m_queryTasks.append(task);

            // 简化日志输出
            qDebug() << QString("  行%1 列%2: RTU=%3, 时间=%4")
                .arg(row).arg(col).arg(rtuId).arg(m_currentTime);
        }
    }
}

bool DayReportParser::executeQueries(QProgressDialog* progress)
{
    if (m_queryTasks.isEmpty()) {
        return true;
    }

    qDebug() << "========== 开始填充数据 ==========";

    // 如果缓存为空，重新查询
    bool cacheReady = true;
    if (m_dataCache.isEmpty()) {
        qWarning() << "缓存为空，重新查询...";
        cacheReady = analyzeAndPrefetch();

        // 即使查询失败，也继续填充（会标记为N/A）
        if (!cacheReady) {
            qWarning() << "查询失败，将所有数据标记为N/A";
        }
    }


    if (progress) {
        progress->setRange(0, m_queryTasks.size());
        progress->setLabelText("正在填充数据...");
    }

    int successCount = 0;
    int failCount = 0;

    for (int i = 0; i < m_queryTasks.size(); ++i) {
        if (progress && progress->wasCanceled()) {
            // 取消时，将剩余任务标记为未执行
            for (int j = i; j < m_queryTasks.size(); ++j) {
                m_queryTasks[j].cell->queryExecuted = false;
            }
            return false;
        }

        const QueryTask& task = m_queryTasks[i];

        QTime time = getTaskTime(task);
        if (!time.isValid()) {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;  // 【保持】标记已执行
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
        // 只有缓存就绪时才尝试查找
        if (cacheReady && findInCache(task.cell->rtuId, timestamp, value)) {
            task.cell->value = QString::number(value, 'f', 2);
            task.cell->queryExecuted = true;
            task.cell->querySuccess = true;
            successCount++;
        }
        else {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;  // 【保持】标记已执行
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

void DayReportParser::restoreToTemplate()
{
    qDebug() << "恢复到模板初始状态...";
    // 清空缓存
    m_dataCache.clear();
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

// ===== 【测试函数】验证数据对应关系 =====
void DayReportParser::runCorrectnessTest()
{
    qDebug() << "";
    qDebug() << "╔════════════════════════════════════════════════════════════╗";
    qDebug() << "║          数据对应关系完整性测试                              ║";
    qDebug() << "╚════════════════════════════════════════════════════════════╝";
    qDebug() << "";

    if (m_queryTasks.isEmpty()) {
        qDebug() << "❌ 无测试数据";
        return;
    }

    // ===== 第1步：选择前3个任务作为测试样本 =====
    int testCount = qMin(3, m_queryTasks.size());
    QList<QueryTask> testTasks;
    for (int i = 0; i < testCount; ++i) {
        testTasks.append(m_queryTasks[i]);
    }

    qDebug() << "【测试样本】选择前" << testCount << "个数据点：";
    for (int i = 0; i < testTasks.size(); ++i) {
        const QueryTask& task = testTasks[i];
        QTime time = getTaskTime(task);
        qDebug() << QString("  样本%1: 行%2列%3, RTU=%4, 时间=%5")
            .arg(i + 1)
            .arg(task.row + 1)
            .arg(task.col + 1)
            .arg(task.cell->rtuId)
            .arg(time.toString("HH:mm:ss"));
    }
    qDebug() << "";

    // ===== 第2步：构造单独的验证查询 =====
    qDebug() << "【验证查询】对这3个数据点单独查询：";

    QHash<QString, QHash<int64_t, float>> verificationData;  // RTU → (时间戳 → 值)

    for (const QueryTask& task : testTasks) {
        QTime time = getTaskTime(task);
        QDateTime dateTime = constructDateTime(m_baseDate, time.toString("HH:mm:ss"));

        // 单独查询这个RTU的这个时间点（前后1分钟）
        QString verifyQuery = QString("%1@%2~%3#60")
            .arg(task.cell->rtuId)
            .arg(dateTime.toString("yyyy-MM-dd HH:mm:ss"))
            .arg(dateTime.addSecs(60).toString("yyyy-MM-dd HH:mm:ss"));

        try {
            auto dataMap = m_fetcher->fetchDataFromAddress(verifyQuery.toStdString());

            if (!dataMap.empty()) {
                auto entry = dataMap.begin();
                int64_t timestamp = entry->first;
                float value = entry->second.empty() ? 0.0f : entry->second[0];

                verificationData[task.cell->rtuId][timestamp] = value;

                // 转换时间戳为可读格式
                QDateTime dt = QDateTime::fromMSecsSinceEpoch(timestamp);

                qDebug() << QString("  ✓ RTU=%1, 时间=%2, 时间戳=%3, 值=%4")
                    .arg(task.cell->rtuId)
                    .arg(dt.toString("HH:mm:ss"))
                    .arg(timestamp)
                    .arg(value, 0, 'f', 2);
            }
            else {
                qDebug() << QString("  ✗ RTU=%1 查询无数据").arg(task.cell->rtuId);
            }
        }
        catch (const std::exception& e) {
            qDebug() << QString("  ✗ RTU=%1 查询失败: %2")
                .arg(task.cell->rtuId)
                .arg(e.what());
        }
    }
    qDebug() << "";

    // ===== 第3步：检查缓存中的数据 =====
    qDebug() << "【缓存检查】查看这3个数据点是否在缓存中：";

    for (const QueryTask& task : testTasks) {
        QTime time = getTaskTime(task);
        QDateTime dateTime = constructDateTime(m_baseDate, time.toString("HH:mm:ss"));
        int64_t timestamp = dateTime.toMSecsSinceEpoch();

        CacheKey key;
        key.rtuId = task.cell->rtuId;
        key.timestamp = timestamp;

        if (m_dataCache.contains(key)) {
            float cachedValue = m_dataCache[key];
            qDebug() << QString("  ✓ 缓存命中: RTU=%1, 时间戳=%2, 缓存值=%3")
                .arg(task.cell->rtuId)
                .arg(timestamp)
                .arg(cachedValue, 0, 'f', 2);
        }
        else {
            qDebug() << QString("  ✗ 缓存未命中: RTU=%1, 时间戳=%2")
                .arg(task.cell->rtuId)
                .arg(timestamp);

            // 尝试模糊查找
            qDebug() << "    → 尝试模糊查找（±60秒）：";
            bool foundNearby = false;
            for (auto it = m_dataCache.constBegin(); it != m_dataCache.constEnd(); ++it) {
                if (it.key().rtuId == task.cell->rtuId) {
                    int64_t diff = qAbs(it.key().timestamp - timestamp);
                    if (diff <= 60000) {  // 60秒
                        QDateTime nearbyDt = QDateTime::fromMSecsSinceEpoch(it.key().timestamp);
                        qDebug() << QString("      找到: 时间戳=%1 (%2), 值=%3, 偏差=%4毫秒")
                            .arg(it.key().timestamp)
                            .arg(nearbyDt.toString("HH:mm:ss"))
                            .arg(it.value(), 0, 'f', 2)
                            .arg(diff);
                        foundNearby = true;
                    }
                }
            }
            if (!foundNearby) {
                qDebug() << "      未找到相近的缓存数据";
            }
        }
    }
    qDebug() << "";

    // ===== 第4步：对比验证值 vs 缓存值 =====
    qDebug() << "【对比验证】单独查询的值 vs 批量查询缓存的值：";

    bool allMatch = true;
    for (const QueryTask& task : testTasks) {
        QTime time = getTaskTime(task);
        QDateTime dateTime = constructDateTime(m_baseDate, time.toString("HH:mm:ss"));
        int64_t timestamp = dateTime.toMSecsSinceEpoch();

        // 从验证查询获取的值
        float verifyValue = 0.0f;
        bool hasVerifyValue = false;
        if (verificationData.contains(task.cell->rtuId)) {
            auto& timeMap = verificationData[task.cell->rtuId];
            if (!timeMap.isEmpty()) {
                verifyValue = timeMap.begin().value();   // 取第一个值
                hasVerifyValue = true;
            }
        }

        // 从缓存获取的值
        float cachedValue = 0.0f;
        bool hasCachedValue = findInCache(task.cell->rtuId, timestamp, cachedValue);

        if (hasVerifyValue && hasCachedValue) {
            float diff = qAbs(verifyValue - cachedValue);
            bool match = diff < 0.01f;  // 允许0.01的误差

            if (match) {
                qDebug() << QString("  ✓ 匹配: RTU=%1, 验证值=%2, 缓存值=%3")
                    .arg(task.cell->rtuId)
                    .arg(verifyValue, 0, 'f', 2)
                    .arg(cachedValue, 0, 'f', 2);
            }
            else {
                qDebug() << QString("  ✗ 不匹配: RTU=%1, 验证值=%2, 缓存值=%3, 差异=%4")
                    .arg(task.cell->rtuId)
                    .arg(verifyValue, 0, 'f', 2)
                    .arg(cachedValue, 0, 'f', 2)
                    .arg(diff, 0, 'f', 4);
                allMatch = false;
            }
        }
        else {
            qDebug() << QString("  ✗ 数据不完整: RTU=%1, 有验证值=%2, 有缓存值=%3")
                .arg(task.cell->rtuId)
                .arg(hasVerifyValue ? "是" : "否")
                .arg(hasCachedValue ? "是" : "否");
            allMatch = false;
        }
    }
    qDebug() << "";

    // ===== 第5步：检查单元格填充结果 =====
    qDebug() << "【单元格检查】查看最终填充到Excel的值：";

    for (const QueryTask& task : testTasks) {
        QString cellValue = task.cell->value.toString();
        bool isExecuted = task.cell->queryExecuted;
        bool isSuccess = task.cell->querySuccess;

        qDebug() << QString("  单元格[%1,%2]: RTU=%3, 显示值=%4, 已执行=%5, 成功=%6")
            .arg(task.row + 1)
            .arg(task.col + 1)
            .arg(task.cell->rtuId)
            .arg(cellValue)
            .arg(isExecuted ? "是" : "否")
            .arg(isSuccess ? "是" : "否");
    }
    qDebug() << "";

    // ===== 最终结论 =====
    qDebug() << "╔════════════════════════════════════════════════════════════╗";
    if (allMatch) {
        qDebug() << "║   测试通过：数据对应关系正确！                            ║";
    }
    else {
        qDebug() << "║   测试失败：发现数据不匹配！                              ║";
    }
    qDebug() << "╚════════════════════════════════════════════════════════════╝";
    qDebug() << "";
}