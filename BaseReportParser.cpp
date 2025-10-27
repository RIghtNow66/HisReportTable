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
    , m_taskWatcher(nullptr) 
    , m_isTaskRunning(false) 
    , m_cancelRequested(0)  
    , m_lastPrefetchSuccessCount(0)  
    , m_lastPrefetchTotalCount(0)
    , m_cacheTimestamp()
    , m_editState(CONFIG_EDIT)
{
    // 创建 FutureWatcher
    m_taskWatcher = new QFutureWatcher<bool>(this);

    // 连接完成信号
    connect(m_taskWatcher, &QFutureWatcher<bool>::finished,
        this, &BaseReportParser::onAsyncTaskFinished);
}

BaseReportParser::~BaseReportParser()
{
    if (m_isTaskRunning) {
        qDebug() << "析构函数：请求停止后台任务...";
        requestCancel();
        // 使用 waitForFinished 确保线程优雅退出
        m_taskFuture.waitForFinished();
        qDebug() << "析构函数：后台任务已停止。";
    }
    delete m_fetcher;
}

void BaseReportParser::startAsyncTask()
{
    if (m_isTaskRunning) {
        qWarning() << "任务已在进行中，无法重复启动。";
        return;
    }

    m_isTaskRunning = true;
    m_cancelRequested.storeRelease(0);

    // 使用 QtConcurrent::run 启动后台任务
    m_taskFuture = QtConcurrent::run([this]() -> bool {
        try {
            // runAsyncTask 是一个虚函数，由子类实现具体任务
            return this->runAsyncTask();
        }
        catch (const std::exception& e) {
            qWarning() << "[后台线程] 发生未捕获的异常：" << e.what();
            // ** [修改] 统一异常处理 **
            emit databaseError(QString("后台任务异常: %1").arg(e.what()));
            return false;
        }
        });

    m_taskWatcher->setFuture(m_taskFuture);
}

void BaseReportParser::onAsyncTaskFinished()
{
    m_isTaskRunning = false;
    bool success = m_taskFuture.result();

    QString message;
    if (m_cancelRequested.loadAcquire()) {
        message = "操作已取消。";
        success = false; // 取消被视为不成功
    }
    else if (success) {
        message = "后台任务成功完成。";
        if (m_model) {
            m_model->markAllCellsClean();
            qDebug() << "预查询完成，已标记所有单元格为干净";
        }
    }
    else {
        message = "后台任务执行失败。";
    }

    // 发射统一的完成信号
    emit asyncTaskCompleted(success, message);
}

void BaseReportParser::requestCancel()
{
    if (m_isTaskRunning) {
        m_cancelRequested.storeRelease(1);
    }
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


// ===== 默认实现（子类可重写） =====

bool BaseReportParser::isTimeMarker(const QString& text) const
{
    return text.startsWith("#t#", Qt::CaseInsensitive);
}

bool BaseReportParser::isDataMarker(const QString& text) const
{
    return text.startsWith("#d#", Qt::CaseInsensitive);
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

BaseReportParser::RescanDiffInfo BaseReportParser::rescanDirtyCells(const QSet<QPoint>& dirtyCells)
{
    qDebug() << "========== 开始增量扫描（增强版） ==========";
    qDebug() << "脏单元格数量：" << dirtyCells.size();

    RescanDiffInfo diffInfo;
    int newCount = 0;
    int modifiedCount = 0;
    int removedCount = 0;
    QSet<int> affectedRows;

    // ===== 阶段1：检测时间标记变化，级联标记同行数据标记 =====
    QSet<QPoint> cascadedDirtyCells = dirtyCells;  // 扩展后的脏单元格集合


    for (const auto& pos : dirtyCells) {
        int row = pos.y();
        int col = pos.x();

        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->displayText();

        // 检测到时间标记或日期标记变化
        if (isTimeMarker(text) || cell->cellType == CellData::TimeMarker ||
            cell->cellType == CellData::DateMarker) {

            qDebug() << QString("检测到时间/日期标记变化：行%1列%2").arg(row).arg(col);
            diffInfo.hasTimeMarkerChange = true;

            // 级联标记同行所有数据标记为脏
            int totalCols = m_model->columnCount();
            for (int c = 0; c < totalCols; ++c) {
                CellData* rowCell = m_model->getCell(row, c);
                if (rowCell && rowCell->cellType == CellData::DataMarker) {
                    QPoint dataPos(c, row);
                    if (!cascadedDirtyCells.contains(dataPos)) {
                        cascadedDirtyCells.insert(dataPos);
                        qDebug() << QString("  级联标记数据标记：行%1列%2").arg(row).arg(c);
                    }
                }
            }
        }
    }

    qDebug() << QString("级联后脏单元格数量：%1").arg(cascadedDirtyCells.size());

    // ===== 【新增】阶段1.5：处理新增的时间标记，设置其 cellType =====
    for (const auto& pos : cascadedDirtyCells) {
        int row = pos.y();
        int col = pos.x();

        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->displayText().trimmed();

        // 识别时间标记但 cellType 未设置的单元格
        if (isTimeMarker(text) && cell->cellType != CellData::TimeMarker) {
            qDebug() << QString("发现未设置类型的时间标记：行%1列%2，文本=%3")
                .arg(row).arg(col).arg(text);

            // 立即设置 cellType
            cell->cellType = CellData::TimeMarker;
            cell->markerText = text;

            // 提取并设置 displayValue
            QString timeValue = extractTime(text);
            if (!timeValue.isEmpty()) {
                cell->displayValue = timeValue;
                qDebug() << QString("  设置时间标记类型：cellType=TimeMarker, displayValue=%1")
                    .arg(timeValue);
            }
        }

        // 识别日期标记但 cellType 未设置的单元格（月报）
        if (cell->cellType != CellData::DateMarker) {
            // 检查是否是日期标记（子类可能需要重写）
            if (text.startsWith("#Date", Qt::CaseInsensitive)) {
                cell->cellType = CellData::DateMarker;
                cell->markerText = text;
                qDebug() << QString("  设置日期标记类型：行%1列%2").arg(row).arg(col);
            }
        }
    }


    // ===== 阶段2：处理所有脏单元格（包括级联的）=====
    for (const auto& pos : cascadedDirtyCells) {
        int row = pos.y();
        int col = pos.x();

        affectedRows.insert(row);

        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->displayText();

        // 【处理数据标记】
        if (isDataMarker(text)) {
            QString rtuId = extractRtuId(text);
            QString oldMarker = m_scannedMarkers.value(pos);

            if (oldMarker.isEmpty()) {
                // ===== 新增数据标记 =====

                // 查找时间上下文
                QString timeStr = findTimeForDataMarker(row, col);
                if (timeStr.isEmpty()) {
                    qWarning() << QString("新增数据标记[%1,%2]时未找到时间上下文，跳过")
                        .arg(row).arg(col);
                    continue;
                }

                m_dataMarkerCells.append({ row, col, rtuId });
                m_scannedMarkers.insert(pos, rtuId);

                QueryTask task;
                task.cell = cell;
                task.row = row;
                task.col = col;
                task.queryPath = findTimeForDataMarker(row, col);
                m_queryTasks.append(task);

                modifiedCount++;
                qDebug() << QString("修改数据标记：行%1列%2，%3 → %4")
                    .arg(row).arg(col).arg(oldMarker).arg(rtuId);
            }
            else {
                // RTU号未变，但可能时间变了（级联导致的）
                // 更新 QueryTask 的时间上下文
                QString newTimeStr = findTimeForDataMarker(row, col);
                for (auto& task : m_queryTasks) {
                    if (task.row == row && task.col == col) {
                        QString oldTimeStr = task.queryPath;
                        if (oldTimeStr != newTimeStr) {
                            // 时间变化，记录到差分
                            QDateTime oldDateTime = constructDateTime(m_baseDate, oldTimeStr);
                            QDateTime newDateTime = constructDateTime(m_baseDate, newTimeStr);

                            if (oldDateTime.isValid() && newDateTime.isValid()) {
                                RescanDiffInfo::ModifiedMarker modifiedMarker;
                                modifiedMarker.rtuId = rtuId;
                                modifiedMarker.oldTimestamp = oldDateTime.toMSecsSinceEpoch();
                                modifiedMarker.newTimestamp = newDateTime.toMSecsSinceEpoch();
                                diffInfo.modifiedMarkers.append(modifiedMarker);

                                qDebug() << QString("时间变化：行%1列%2，%3 → %4")
                                    .arg(row).arg(col).arg(oldTimeStr).arg(newTimeStr);
                            }

                            task.queryPath = newTimeStr;  // 更新时间
                        }
                        break;
                    }
                }
            }
        }
        // 【移除数据标记】
        else {
            if (m_scannedMarkers.contains(pos)) {
                QString oldRtuId = m_scannedMarkers.value(pos);

                // 计算旧时间戳用于缓存清理
                int64_t oldTimestamp = calculateTimestampForMarker(row, col);
                if (oldTimestamp > 0) {
                    RescanDiffInfo::RemovedMarker removedMarker;
                    removedMarker.rtuId = oldRtuId;
                    removedMarker.timestamp = oldTimestamp;
                    diffInfo.removedMarkers.append(removedMarker);
                }

                // 移除记录
                for (auto it = m_dataMarkerCells.begin(); it != m_dataMarkerCells.end(); ) {
                    if (it->row == row && it->col == col) {
                        it = m_dataMarkerCells.erase(it);
                    }
                    else {

                        ++it;
                    }
                }

                for (auto it = m_queryTasks.begin(); it != m_queryTasks.end(); ) {
                    if (it->row == row && it->col == col) {
                        it = m_queryTasks.erase(it);
                    }
                    else {

                        ++it;
                    }
                }

                m_scannedMarkers.remove(pos);
                removedCount++;
                qDebug() << QString("移除数据标记：行%1列%2，RTU=%3")
                    .arg(row).arg(col).arg(oldRtuId);
            }
        }
    }

    qDebug() << QString("增量扫描完成：新增 %1，修改 %2，移除 %3")
        .arg(newCount).arg(modifiedCount).arg(removedCount);
    qDebug() << QString("受影响的行数：%1").arg(affectedRows.size());
    qDebug() << QString("当前查询任务数：%1").arg(m_queryTasks.size());
    qDebug() << QString("差分信息：新增=%1，删除=%2，时间修改=%3")
        .arg(diffInfo.newMarkerCount)
        .arg(diffInfo.removedMarkers.size())
        .arg(diffInfo.modifiedMarkers.size());

    // 通知子类
    onRescanCompleted(newCount, modifiedCount, removedCount, affectedRows);

    qDebug() << "========================================";

    return diffInfo;
}

// ===== 计算数据标记对应的时间戳 =====
int64_t BaseReportParser::calculateTimestampForMarker(int row, int col)
{
    // 查找该数据标记的时间上下文
    QString timeStr = findTimeForDataMarker(row, col);
    if (timeStr.isEmpty()) {
        qWarning() << QString("无法计算时间戳：行%1列%2无时间上下文").arg(row).arg(col);
        return -1;
    }

    // 构造完整的日期时间
    QDateTime dateTime = constructDateTime(m_baseDate, timeStr);
    if (!dateTime.isValid()) {
        qWarning() << QString("无法计算时间戳：日期时间无效，日期=%1，时间=%2")
            .arg(m_baseDate).arg(timeStr);
        return -1;
    }

    return dateTime.toMSecsSinceEpoch();
}

// ===== 根据差分信息清理缓存 =====
void BaseReportParser::cleanupCacheByDiff(const RescanDiffInfo& diffInfo)
{
    if (diffInfo.removedMarkers.isEmpty() && diffInfo.modifiedMarkers.isEmpty()) {
        qDebug() << "差分为空，无需清理缓存";
        return;
    }

    qDebug() << "========== 开始缓存清理（差分方式） ==========";

    QMutexLocker locker(&m_cacheMutex);

    int cleanedCount = 0;

    // 清理被删除的标记
    for (const auto& removed : diffInfo.removedMarkers) {
        CacheKey key;
        key.rtuId = removed.rtuId;
        key.timestamp = removed.timestamp;

        if (m_dataCache.contains(key)) {
            m_dataCache.remove(key);
            cleanedCount++;
            qDebug() << QString("  删除缓存：RTU=%1，时间戳=%2")
                .arg(removed.rtuId).arg(removed.timestamp);
        }

        // 同时清理索引缓存
        if (m_rtuidIndexCache.contains(removed.rtuId)) {
            auto& indexList = m_rtuidIndexCache[removed.rtuId];
            indexList.erase(
                std::remove_if(indexList.begin(), indexList.end(),
                    [&](const QPair<int64_t, float>& pair) {
                        return pair.first == removed.timestamp;
                    }),
                indexList.end()
            );
        }
    }

    // 清理时间修改的标记（删除旧时间戳）
    for (const auto& modified : diffInfo.modifiedMarkers) {
        CacheKey oldKey;
        oldKey.rtuId = modified.rtuId;
        oldKey.timestamp = modified.oldTimestamp;

        if (m_dataCache.contains(oldKey)) {
            m_dataCache.remove(oldKey);
            cleanedCount++;
            qDebug() << QString("  删除旧缓存：RTU=%1，旧时间戳=%2，新时间戳=%3")
                .arg(modified.rtuId)
                .arg(modified.oldTimestamp)
                .arg(modified.newTimestamp);
        }

        // 同时清理索引缓存中的旧时间戳
        if (m_rtuidIndexCache.contains(modified.rtuId)) {
            auto& indexList = m_rtuidIndexCache[modified.rtuId];
            indexList.erase(
                std::remove_if(indexList.begin(), indexList.end(),
                    [&](const QPair<int64_t, float>& pair) {
                        return pair.first == modified.oldTimestamp;
                    }),
                indexList.end()
            );
        }
    }

    qDebug() << QString("缓存清理完成：清理了 %1 个缓存项").arg(cleanedCount);
    qDebug() << QString("当前缓存大小：%1 项").arg(m_dataCache.size());
    qDebug() << "============================================";
}