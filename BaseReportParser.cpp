#include "BaseReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"

#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept>
#include <QProgressDialog>
#include <QtConcurrent>
#include <QEventLoop>
#include <QTimer>

BaseReportParser::BaseReportParser(ReportDataModel* model, QObject* parent)
    : QObject(parent)
    , m_model(model)
    , m_fetcher(new TaosDataFetcher())
    , m_dateFound(false)
    , m_prefetchWatcher(nullptr)
    , m_isPrefetching(false)
    , m_stopRequested(0)
    , m_lastPrefetchSuccessCount(0)  
    , m_lastPrefetchTotalCount(0)
    , m_cacheTimestamp()
    , m_editState(CONFIG_EDIT)
{
    // 创建 FutureWatcher
    m_prefetchWatcher = new QFutureWatcher<bool>(this);

    // 连接完成信号
    connect(m_prefetchWatcher, &QFutureWatcher<bool>::finished,
        this, &BaseReportParser::onPrefetchFinished);
}

BaseReportParser::~BaseReportParser()
{
    if (m_prefetchWatcher) {
        disconnect(m_prefetchWatcher, nullptr, this, nullptr);
    }

    if (m_isPrefetching) {
        qDebug() << "请求停止预查询...";
        m_stopRequested.storeRelease(1);

        if (!m_prefetchFuture.isFinished()) {
            qDebug() << "等待后台线程退出（最多3秒）...";

            // 修改：使用 waitForFinished 替代事件循环
            m_prefetchFuture.waitForFinished();
        }
    }
    delete m_fetcher;
}

void BaseReportParser::onPrefetchFinished()
{
    m_isPrefetching = false;

    bool success = m_prefetchFuture.result();

    QMutexLocker locker(&m_cacheMutex);
    int dataCount = m_dataCache.size();
    locker.unlock();

    // 获取统计信息
    int successCount = m_lastPrefetchSuccessCount;
    int totalCount = m_lastPrefetchTotalCount;

    setEditState(CONFIG_EDIT);
    m_model->updateEditability();

    if (dataCount > 0) {
        qDebug() << QString("预查询完成：已缓存 %1 个数据点（成功 %2/%3）")
            .arg(dataCount).arg(successCount).arg(totalCount);
        emit prefetchCompleted(true, dataCount, successCount, totalCount);
    }
    else {
        qWarning() << QString("预查询失败：无数据（成功 %1/%2）")
            .arg(successCount).arg(totalCount);
        emit prefetchCompleted(false, 0, successCount, totalCount);
    }

    // 重置统计信息
    m_lastPrefetchSuccessCount = 0;
    m_lastPrefetchTotalCount = 0;

}

void BaseReportParser::clearCache()
{
    QMutexLocker locker(&m_cacheMutex);
    qDebug() << "清空缓存：" << m_dataCache.size() << "个数据点";
    m_dataCache.clear();
    m_rtuidIndexCache.clear();  // 新增
    m_cacheTimestamp = QDateTime();
}

bool BaseReportParser::findInCache(const QString& rtuId, int64_t timestamp, float& value)
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

    // 2. 容错匹配：使用索引加速
    const int64_t tolerance = 300000;

    if (!m_rtuidIndexCache.contains(rtuId)) {
        return false;
    }

    const QList<QPair<int64_t, float>>& timeValues = m_rtuidIndexCache[rtuId];

    int64_t closestDiff = std::numeric_limits<int64_t>::max();
    bool found = false;
    float closestValue = 0.0f;

    for (const auto& pair : timeValues) {
        int64_t diff = qAbs(pair.first - timestamp);
        if (diff <= tolerance && diff < closestDiff) {
            closestDiff = diff;
            closestValue = pair.second;
            found = true;
        }
    }

    if (found) {
        value = closestValue;
        return true;
    }

    return false;
}

bool BaseReportParser::executeSingleQuery(const QString& rtuList,
    const QTime& startTime,
    const QTime& endTime,
    int intervalSeconds)
{
    QString query;
    QString startDateStr, endDateStr;

    if (getDateRange(startDateStr, endDateStr)) {
        // 月报模式
        query = QString("%1@%2 %3~%4 %5#%6")
            .arg(rtuList)
            .arg(startDateStr)
            .arg(startTime.toString("HH:mm:ss"))
            .arg(endDateStr)
            .arg(endTime.toString("HH:mm:ss"))
            .arg(intervalSeconds);
    }
    else {
        // 日报模式
        query = QString("%1@%2 %3~%2 %4#%5")
            .arg(rtuList)
            .arg(m_baseDate)
            .arg(startTime.toString("HH:mm:ss"))
            .arg(endTime.addSecs(60).toString("HH:mm:ss"))
            .arg(intervalSeconds);
    }

    qDebug() << "  查询地址：" << query;

    auto startQueryTime = QDateTime::currentDateTime();  

    try {
        auto dataMap = m_fetcher->fetchDataFromAddress(query.toStdString());

        auto endQueryTime = QDateTime::currentDateTime();  
        qDebug() << "  数据库查询耗时：" << startQueryTime.msecsTo(endQueryTime) << "ms";  
        qDebug() << "  返回时间点数量：" << dataMap.size();

        qDebug() << "  === 返回的所有时间戳 ===";
        for (auto it = dataMap.begin(); it != dataMap.end(); ++it) {
            int64_t ts = it->first;
            QDateTime dt = QDateTime::fromMSecsSinceEpoch(ts);
            qDebug() << QString("    时间戳: %1 -> %2").arg(ts).arg(dt.toString("yyyy-MM-dd HH:mm:ss"));
        }
        qDebug() << "  =========================";

        if (dataMap.empty()) {
            qWarning() << "  查询无数据";
            emit databaseError("未获取到有效数据，请检查TDengine连接");
            return false;
        }

        // 【优化】在锁外准备数据
        QHash<CacheKey, float> tempCache;
        QHash<QString, QList<QPair<int64_t, float>>> tempIndexCache;
        QStringList rtuArray = rtuList.split(",");

        for (auto it = dataMap.begin(); it != dataMap.end(); ++it) {
            int64_t timestamp = it->first;
            const std::vector<float>& values = it->second;

            for (int i = 0; i < rtuArray.size() && i < values.size(); ++i) {
                CacheKey key;
                key.rtuId = rtuArray[i];
                key.timestamp = timestamp;
                tempCache[key] = values[i];

                tempIndexCache[rtuArray[i]].append(qMakePair(timestamp, values[i]));
            }
        }

        // 【优化】快速持锁写入
        {
            QMutexLocker locker(&m_cacheMutex);
            m_dataCache.unite(tempCache);

            for (auto it = tempIndexCache.constBegin(); it != tempIndexCache.constEnd(); ++it) {
                m_rtuidIndexCache[it.key()].append(it.value());
            }

            m_cacheTimestamp = QDateTime::currentDateTime();
        }

        return true;
    }
    catch (const std::exception& e) {
        qWarning() << "  查询失败：" << e.what();
        emit databaseError(QString("数据查询失败: %1").arg(e.what()));
        return false;
    }
}

// 虚函数：获取日期范围（月报重写）
bool BaseReportParser::getDateRange(QString& startDate, QString& endDate)
{
    // 基类默认返回 false（日报不需要日期范围）
    startDate.clear();
    endDate.clear();
    return false;
}

bool BaseReportParser::shouldMergeBlocks(const TimeBlock& block1, const TimeBlock& block2)
{
    // 计算两个块之间的间隔
    int gapMinutes = block1.endTime.secsTo(block2.startTime) / 60;

    // 经验阈值：间隔超过2小时（120分钟），就不合并
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

QList<BaseReportParser::TimeBlock> BaseReportParser::identifyTimeBlocks()
{
    qDebug() << "【月报】identifyTimeBlocks() 被调用";
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

bool BaseReportParser::analyzeAndPrefetch()
{
    // 1. 识别时间块
    QList<TimeBlock> blocks = identifyTimeBlocks();

    if (m_stopRequested.loadAcquire()) return false;

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
        if (m_stopRequested.loadAcquire()) {
            qDebug() << "后台查询被中断";
            // 记录统计信息
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


        emit prefetchProgress(i + 1, mergedBlocks.size());

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

// ===== 默认实现（子类可重写） =====

bool BaseReportParser::isTimeMarker(const QString& text) const
{
    return text.startsWith("#t#", Qt::CaseInsensitive);
}

bool BaseReportParser::isDataMarker(const QString& text) const
{
    return text.startsWith("#d#", Qt::CaseInsensitive);
}

QString BaseReportParser::extractTime(const QString& text)
{
    // 默认实现：#t#0:00 → "00:00:00"
    QString timeStr = text.mid(3).trimmed();
    QStringList parts = timeStr.split(":");

    if (parts.size() == 2) {
        timeStr += ":00";
    }
    else if (parts.size() != 3) {
        qWarning() << "时间格式错误:" << text;
        return "00:00:00";
    }

    QTime time = QTime::fromString(timeStr, "H:mm:ss");
    if (!time.isValid()) {
        qWarning() << "时间解析失败:" << timeStr;
        return "00:00:00";
    }

    return time.toString("HH:mm:ss");
}

QString BaseReportParser::extractRtuId(const QString& text)
{
    // #d#AIRTU034700019 → "AIRTU034700019"
    return text.mid(3).trimmed();
}

void BaseReportParser::runCorrectnessTest()
{
    qDebug() << "基类测试函数 - 子类应该重写此方法";
}

bool BaseReportParser::isCacheValid() const
{
    if (m_dataCache.isEmpty()) {
        return false;
    }

    if (!m_cacheTimestamp.isValid()) {
        return false;
    }

    QDateTime now = QDateTime::currentDateTime();
    qint64 hours = m_cacheTimestamp.secsTo(now) / 3600;

    bool valid = hours < CACHE_EXPIRE_HOURS;

    if (!valid) {
        qDebug() << QString("缓存已过期: %1 小时前创建").arg(hours);
    }

    return valid;
}

void BaseReportParser::invalidateCache()
{
    qDebug() << "使缓存失效";
    clearCache();
}