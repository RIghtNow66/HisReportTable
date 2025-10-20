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

            // 使用 QDeadlineTimer 实现超时
            QDeadlineTimer deadline(3000);
            while (!m_prefetchFuture.isFinished() && !deadline.hasExpired()) {
                QThread::msleep(50);  // 每50ms检查一次
                QCoreApplication::processEvents();  // 处理信号
            }

            if (!m_prefetchFuture.isFinished()) {
                qWarning() << "后台线程未能及时退出，强制析构";
            }
        }
    }
    delete m_fetcher;
}

void BaseReportParser::onPrefetchFinished()
{
    m_isPrefetching = false;

    bool success = m_prefetchFuture.result();

    if (success) {
        QMutexLocker locker(&m_cacheMutex);
        qDebug() << "预查询完成：已缓存" << m_dataCache.size() << "个数据点";
    }
    else {
        qWarning() << "预查询失败";
    }
}

void BaseReportParser::clearCache()
{
    QMutexLocker locker(&m_cacheMutex);
    qDebug() << "清空缓存：" << m_dataCache.size() << "个数据点";
    m_dataCache.clear();
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

    // 2. 容错匹配：±60秒内的最近时间点
    const int64_t tolerance = 60000;  // 60秒（毫秒）

    int64_t closestDiff = std::numeric_limits<int64_t>::max();
    bool found = false;
    float closestValue = 0.0f;

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

bool BaseReportParser::executeSingleQuery(const QString& rtuList,
    const QTime& startTime,
    const QTime& endTime,
    int intervalSeconds)
{
    // 【修改】支持日期范围查询（月报）
    QString query;

    // 检查是否为月报查询（子类会传入完整的 TimeBlock）
    // 这里需要通过虚函数获取日期信息
    QString startDateStr, endDateStr;
    if (getDateRange(startDateStr, endDateStr)) {
        // 月报模式：构建带日期的查询
        query = QString("%1@%2 %3~%4 %5#%6")
            .arg(rtuList)
            .arg(startDateStr)                          // 2025-07-10
            .arg(startTime.toString("HH:mm:ss"))        // 08:30:00
            .arg(endDateStr)                            // 2025-07-20
            .arg(endTime.toString("HH:mm:ss"))          // 08:31:00
            .arg(intervalSeconds);
    }
    else {
        // 日报模式：原有逻辑
        query = QString("%1@%2 %3~%2 %4#%5")
            .arg(rtuList)
            .arg(m_baseDate)
            .arg(startTime.toString("HH:mm:ss"))
            .arg(endTime.addSecs(60).toString("HH:mm:ss"))
            .arg(intervalSeconds);
    }

    qDebug() << "  查询地址：" << query;

    try {
        // 执行查询
        auto dataMap = m_fetcher->fetchDataFromAddress(query.toStdString());

        qDebug() << "  返回时间点数量：" << dataMap.size();

        if (dataMap.empty()) {
            qWarning() << "  查询无数据";
            emit databaseError("未获取到有效数据，请检查TDengine连接");
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

            qDebug() << QString("已缓存 %1 个时间点 × %2 个RTU = %3 个值")
                .arg(dataMap.size())
                .arg(rtuArray.size())
                .arg(m_dataCache.size());
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
        mergedBlocks.append(blocks[0]);
    }
    else {
        TimeBlock currentMerged = blocks[0];

        for (int i = 1; i < blocks.size(); ++i) {
            if (shouldMergeBlocks(currentMerged, blocks[i])) {
                qDebug() << QString("  合并块 %1~%2 和 %3~%4")
                    .arg(currentMerged.startTime.toString("HH:mm"))
                    .arg(currentMerged.endTime.toString("HH:mm"))
                    .arg(blocks[i].startTime.toString("HH:mm"))
                    .arg(blocks[i].endTime.toString("HH:mm"));

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
    bool allSuccess = true;
    int intervalSeconds = getQueryIntervalSeconds();

    for (int i = 0; i < mergedBlocks.size(); ++i) {
        if (m_stopRequested.loadAcquire()) {
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
            intervalSeconds);

        if (!success) {
            qWarning() << "查询失败";
            allSuccess = false;
        }
    }

    return allSuccess;
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