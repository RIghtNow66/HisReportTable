#pragma once
#ifndef DAYREPORTPARSER_H
#define DAYREPORTPARSER_H

#include <QObject>
#include <QString>
#include <QDateTime>
#include <QList>

class ReportDataModel;
class TaosDataFetcher;
struct CellData;

/**
 * @brief 日报解析器
 *
 * 负责解析 ##Day_xxx.xlsx 格式的日报文件
 * 识别 #Date、#t#、#d# 标记，构建查询地址并执行查询
 */
class DayReportParser : public QObject
{
    Q_OBJECT

public:
    explicit DayReportParser(ReportDataModel* model, QObject* parent = nullptr);
    ~DayReportParser();

    // ===== 核心接口 =====

    /**
     * @brief 扫描并解析所有单元格标记
     * @return 是否成功（至少找到#Date且有待查询单元格）
     */
    bool scanAndParse();

    /**
     * @brief 执行所有查询任务
     * @return 是否全部成功
     */
    bool executeQueries();

    // ===== 状态查询 =====

    bool isValid() const { return m_dateFound; }
    QString getBaseDate() const { return m_baseDate; }
    int getPendingQueryCount() const { return m_queryTasks.size(); }

signals:
    void parseProgress(int current, int total);      // 解析进度
    void queryProgress(int current, int total);      // 查询进度
    void parseCompleted(bool success, QString message);
    void queryCompleted(int successCount, int failCount);

private:
    // ===== 第一阶段：标记解析 =====

    /**
     * @brief 查找并解析 #Date 标记
     * @return 是否找到有效的#Date
     */
    bool findDateMarker();

    /**
     * @brief 解析一行中的所有标记
     * @param row 行号
     */
    void parseRow(int row);

    // ===== 第二阶段：查询执行 =====

    /**
     * @brief 执行单个查询任务
     * @param queryPath 查询地址
     * @return 查询结果（单值）
     */
    QVariant querySinglePoint(const QString& queryPath);

    // ===== 工具函数 =====

    bool isDateMarker(const QString& text) const;
    bool isTimeMarker(const QString& text) const;
    bool isDataMarker(const QString& text) const;

    QString extractDate(const QString& text);        // 从 "#Date:2024-01-01" 提取
    QString extractTime(const QString& text);        // 从 "#t#0:00" 提取并格式化
    QString extractRtuId(const QString& text);       // 从 "#d#AIRTU..." 提取

    /**
     * @brief 构造查询地址
     * @param rtuId RTU号
     * @param baseTime 起始时间
     * @return 查询地址字符串
     */
    QString buildQueryPath(const QString& rtuId, const QDateTime& baseTime);

    /**
     * @brief 拼接日期和时间
     * @param date "2024-01-01"
     * @param time "00:00:00"
     * @return QDateTime对象
     */
    QDateTime constructDateTime(const QString& date, const QString& time);

private:
    // ===== 数据成员 =====

    ReportDataModel* m_model;              // 数据模型指针
    TaosDataFetcher* m_fetcher;            // 数据查询器

    // 解析状态
    QString m_baseDate;                    // 基准日期（从 #Date 提取，如 "2024-01-01"）
    QString m_currentTime;                 // 当前有效时间（从最近的 #t# 提取，如 "00:00:00"）
    bool m_dateFound;                      // 是否找到 #Date 标记

    // 查询任务
    struct QueryTask {
        CellData* cell;                    // 单元格指针
        int row;                           // 行号
        int col;                           // 列号
        QString queryPath;                 // 查询地址
    };

    QList<QueryTask> m_queryTasks;         // 待执行的查询任务列表
};

#endif // DAYREPORTPARSER_H
