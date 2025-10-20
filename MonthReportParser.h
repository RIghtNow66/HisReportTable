#pragma once

#pragma once
#ifndef MONTHREPORTPARSER_H
#define MONTHREPORTPARSER_H

#include "BaseReportParser.h"

/**
 * @brief 月报解析器
 * 处理格式：#Date1:2024-01, #Date2:08:30, #t#1, #d#RTU号
 * 文件命名：##Month_xxx.xlsx
 */
class MonthReportParser : public BaseReportParser
{
    Q_OBJECT

public:
    explicit MonthReportParser(ReportDataModel* model, QObject* parent = nullptr);
    ~MonthReportParser() override = default;

    // ===== 实现基类接口 =====
    bool scanAndParse() override;
    bool executeQueries(QProgressDialog* progress) override;
    void restoreToTemplate() override;
    void runCorrectnessTest() override;

    // ===== 月报特有接口 =====
    QString getBaseYearMonth() const { return m_baseYearMonth; }
    QString getBaseTime() const { return m_baseTime; }

protected:
    // ===== 实现纯虚函数 =====
    bool findDateMarker() override;
    void parseRow(int row) override;
    QTime getTaskTime(const QueryTask& task) override;
    QDateTime constructDateTime(const QString& date, const QString& time) override;
    int getQueryIntervalSeconds() const override { return 60; }  // 月报间隔24小时

    QList<TimeBlock> identifyTimeBlocks() override;
    bool getDateRange(QString& startDate, QString& endDate) override;

    // ===== 重写标记识别 =====
    QString extractTime(const QString& text) override;

    // ==== = 重写预查询逻辑 ==== =
    bool analyzeAndPrefetch() override;

private:
    // ===== 月报特有的标记识别 =====
    bool isDate1Marker(const QString& text) const;
    bool isDate2Marker(const QString& text) const;
    QString extractYearMonth(const QString& text);
    QString extractTimeOfDay(const QString& text);
    int extractDay(const QString& text);

    void collectActualDays();
    int getMinDay() const;
    int getMaxDay() const;

private:
    QString m_baseYearMonth;   // "2024-01"
    QString m_baseTime;        // "08:30:00"

    QSet<int> m_actualDays; // 实际出现的日期集合：{ 10, 11, 12, ..., 20 }

    // 临时存储当前查询的日期范围（用于 getDateRange）
    QString m_currentQueryStartDate;
    QString m_currentQueryEndDate;
};

#endif // MONTHREPORTPARSER_H
