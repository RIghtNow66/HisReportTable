#pragma once
#ifndef DAYREPORTPARSER_H
#define DAYREPORTPARSER_H

#include <QObject>
#include <QString>
#include <QDateTime>
#include <QList>
#include <QHash>
#include <QTime>
#include <QThread>
#include <QFuture>
#include <QFutureWatcher> 

class ReportDataModel;
class TaosDataFetcher;
struct CellData;
class QProgressDialog;


class DayReportParser : public QObject
{
    Q_OBJECT

public:
    explicit DayReportParser(ReportDataModel* model, QObject* parent = nullptr);
    ~DayReportParser();

    // ===== 核心接口 =====
    bool scanAndParse();
    bool executeQueries(QProgressDialog* progress);
    void restoreToTemplate();
    void clearCache();

    // ===== 状态查询 =====
    bool isValid() const { return m_dateFound; }
    QString getBaseDate() const { return m_baseDate; }
    int getPendingQueryCount() const { return m_queryTasks.size(); }
    bool isPrefetching() const { return m_isPrefetching; }  

    void runCorrectnessTest();

signals:
    void parseProgress(int current, int total);
    void queryProgress(int current, int total);
    void parseCompleted(bool success, QString message);
    void queryCompleted(int successCount, int failCount);

public:
    // 缓存键
    struct CacheKey {
        QString rtuId;
        int64_t timestamp;

        bool operator==(const CacheKey& other) const {
            return rtuId == other.rtuId && timestamp == other.timestamp;
        }
    };

private:
    // 时间块结构
    struct TimeBlock {
        QTime startTime;
        QTime endTime;
        QList<int> taskIndices;  // 属于这个块的任务索引

        bool isValid() const {
            return startTime.isValid() && endTime.isValid();
        }
    };

    // 缓存：(RTU, 时间戳) → 值
    QHash<CacheKey, float> m_dataCache;

    // ===== 数据成员 =====
    ReportDataModel* m_model;
    TaosDataFetcher* m_fetcher;

    // 解析状态
    QString m_baseDate;
    QString m_currentTime;
    bool m_dateFound;

    // 查询任务结构体
    struct QueryTask {
        CellData* cell;
        int row;
        int col;
        QString queryPath;
    };

    QList<QueryTask> m_queryTasks;

	// 预查询相关
    QFuture<bool> m_prefetchFuture;
    QFutureWatcher<bool>* m_prefetchWatcher;
    bool m_isPrefetching;

private slots:
    void onPrefetchFinished();

private:
    // ===== 内部处理函数 =====
    bool findDateMarker();
    void parseRow(int row);
    QVariant querySinglePoint(const QString& queryPath);

    // ===== 工具函数 =====
    bool isDateMarker(const QString& text) const;
    bool isTimeMarker(const QString& text) const;
    bool isDataMarker(const QString& text) const;

    QString extractDate(const QString& text);
    QString extractTime(const QString& text);
    QString extractRtuId(const QString& text);

    QString buildQueryPath(const QString& rtuId, const QDateTime& baseTime);
    QDateTime constructDateTime(const QString& date, const QString& time);

    // =========智能决策查询算法相关函数==========
    // 智能分析并预查询
    bool analyzeAndPrefetch();

    // 识别连续的时间块
    QList<TimeBlock> identifyTimeBlocks();

    // 决策是否合并查询
    bool shouldMergeBlocks(const TimeBlock& block1, const TimeBlock& block2);

    // 执行单次查询
    bool executeSingleQuery(const QString& rtuList,
        const QTime& startTime,
        const QTime& endTime,
        int intervalSeconds);

    // 从查询路径提取时间
    QTime DayReportParser::getTaskTime(const QueryTask& task);

    // 从缓存查找值
    bool findInCache(const QString& rtuId, int64_t timestamp, float& value);

};

inline uint qHash(const DayReportParser::CacheKey& key, uint seed = 0) {
    return qHash(key.rtuId, seed) ^ qHash(key.timestamp, seed);
}

#endif // DAYREPORTPARSER_H