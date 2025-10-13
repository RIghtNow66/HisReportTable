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

    // ===== 状态查询 =====
    bool isValid() const { return m_dateFound; }
    QString getBaseDate() const { return m_baseDate; }
    int getPendingQueryCount() const { return m_queryTasks.size(); }

signals:
    void parseProgress(int current, int total);
    void queryProgress(int current, int total);
    void parseCompleted(bool success, QString message);
    void queryCompleted(int successCount, int failCount);

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

private:
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
};

#endif // DAYREPORTPARSER_H