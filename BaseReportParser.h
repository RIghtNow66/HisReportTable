#pragma once
#ifndef BASEREPORTPARSER_H
#define BASEREPORTPARSER_H

#include <QObject>
#include <QString>
#include <QDateTime>
#include <QList>
#include <QHash>
#include <QTime>
#include <QMutex>
#include <QFuture>
#include <QFutureWatcher>
#include <QAtomicInt>

class ReportDataModel;
class TaosDataFetcher;
struct CellData;
class QProgressDialog;

/**
 * @brief 报表解析器基类
 * 提供通用的解析流程和缓存机制
 */
class BaseReportParser : public QObject
{
    Q_OBJECT

public:
    // ===== 缓存键结构（所有报表通用） =====
    struct CacheKey {
        QString rtuId;
        int64_t timestamp;

        bool operator==(const CacheKey& other) const {
            return rtuId == other.rtuId && timestamp == other.timestamp;
        }
    };

    // ===== 查询任务结构（所有报表通用） =====
    struct QueryTask {
        CellData* cell;
        int row;
        int col;
        QString queryPath;  // 可选，某些子类可能不需要
    };

    // ===== 时间块结构（用于智能查询优化） =====
    struct TimeBlock {
        QTime startTime;
        QTime endTime;
        QString startDate;  // 用于月报：如 "2025-07-10"
        QString endDate;    // 用于月报：如 "2025-07-20"
        QList<int> taskIndices;

        bool isValid() const {
            return startTime.isValid() && endTime.isValid();
        }

        // 判断是否为日期范围查询（月报）
        bool isDateRange() const {
            return !startDate.isEmpty() && !endDate.isEmpty();
        }
    };

public:
    explicit BaseReportParser(ReportDataModel* model, QObject* parent = nullptr);
    virtual ~BaseReportParser();

    // ===== 核心接口（子类必须实现） =====
    virtual bool scanAndParse() = 0;                     // 扫描并解析模板
    virtual bool executeQueries(QProgressDialog* progress) = 0;  // 执行查询
    virtual void restoreToTemplate() = 0;                // 恢复到模板状态

    // ===== 通用接口 =====
    void clearCache();                                   // 清空缓存
    bool isValid() const { return m_dateFound; }
    int getPendingQueryCount() const { return m_queryTasks.size(); }
    bool isPrefetching() const { return m_isPrefetching; }

    // ===== 测试接口 =====
    virtual void runCorrectnessTest();                   // 验证数据对应关系

signals:
    void parseProgress(int current, int total);
    void queryProgress(int current, int total);
    void parseCompleted(bool success, QString message);
    void queryCompleted(int successCount, int failCount);
    void prefetchProgress(int current, int total);
    void databaseError(QString errorMessage);

    void prefetchCompleted(bool hasData, int dataCount, int successCount, int totalCount);

protected:
    // ===== 子类需要实现的纯虚函数 =====

    /**
     * @brief 查找日期标记（不同报表的日期标记格式不同）
     * @return 是否找到有效的日期标记
     */
    virtual bool findDateMarker() = 0;

    /**
     * @brief 解析单行数据（识别时间标记和数据标记）
     * @param row 行号
     */
    virtual void parseRow(int row) = 0;

    /**
     * @brief 从任务中提取时间（不同报表的时间提取方式不同）
     * @param task 查询任务
     * @return 时间对象
     */
    virtual QTime getTaskTime(const QueryTask& task) = 0;

    /**
     * @brief 构造完整的日期时间（不同报表的组合规则不同）
     * @param date 日期字符串
     * @param time 时间字符串
     * @return 日期时间对象
     */
    virtual QDateTime constructDateTime(const QString& date, const QString& time) = 0;

    /**
     * @brief 获取查询间隔秒数（日报60秒，月报86400秒）
     * @return 间隔秒数
     */
    virtual int getQueryIntervalSeconds() const = 0;

    // ===== 通用工具函数（子类可直接使用） =====

    /**
     * @brief 从缓存中查找数据
     * @param rtuId RTU编号
     * @param timestamp 时间戳（毫秒）
     * @param value 输出参数，找到的值
     * @return 是否找到
     */
    bool findInCache(const QString& rtuId, int64_t timestamp, float& value);

    /**
     * @brief 执行单次查询
     * @param rtuList RTU列表（逗号分隔）
     * @param startTime 起始时间
     * @param endTime 结束时间
     * @param intervalSeconds 间隔秒数
     * @return 是否成功
     */
    bool executeSingleQuery(const QString& rtuList,
        const QTime& startTime,
        const QTime& endTime,
        int intervalSeconds);

    /**
     * @brief 智能分析并预查询
     * @return 是否成功
     */
    virtual bool analyzeAndPrefetch();

    /**
     * @brief 识别连续的时间块
     * @return 时间块列表
     */
    virtual QList<TimeBlock> identifyTimeBlocks();

    /**
     * @brief 判断两个时间块是否应该合并
     * @param block1 时间块1
     * @param block2 时间块2
     * @return 是否应该合并
     */
    bool shouldMergeBlocks(const TimeBlock& block1, const TimeBlock& block2);

    /**
    * @brief 获取日期范围（月报需要重写）
    * @param startDate 输出：起始日期（如 "2025-07-10"）
    * @param endDate 输出：结束日期（如 "2025-07-20"）
    * @return 是否为日期范围查询
    */
    virtual bool getDateRange(QString& startDate, QString& endDate);

    // ===== 标记识别（子类可选择性重写） =====
    virtual bool isTimeMarker(const QString& text) const;
    virtual bool isDataMarker(const QString& text) const;

    virtual QString extractTime(const QString& text);
    virtual QString extractRtuId(const QString& text);

protected slots:
    void onPrefetchFinished();

protected:
    // ===== 数据成员 =====
    ReportDataModel* m_model;          // 数据模型
    TaosDataFetcher* m_fetcher;        // 数据查询器

    // 解析状态
    bool m_dateFound;                  // 是否找到日期标记
    QString m_baseDate;                // 基准日期（格式由子类决定）
    QString m_currentTime;             // 当前时间

    // 查询任务
    QList<QueryTask> m_queryTasks;     // 待查询任务列表

    // 缓存
    QHash<CacheKey, float> m_dataCache;  // 数据缓存
    QMutex m_cacheMutex;               // 缓存互斥锁

    // 预查询
    QFuture<bool> m_prefetchFuture;
    QFutureWatcher<bool>* m_prefetchWatcher;
    bool m_isPrefetching;
    QAtomicInt m_stopRequested;

    int m_lastPrefetchSuccessCount;
    int m_lastPrefetchTotalCount;
};

// Hash 函数（用于 QHash）
inline uint qHash(const BaseReportParser::CacheKey& key, uint seed = 0) {
    return qHash(key.rtuId, seed) ^ qHash(key.timestamp, seed);
}

#endif // BASEREPORTPARSER_H
