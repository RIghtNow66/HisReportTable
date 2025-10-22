#pragma once
#ifndef DAYREPORTPARSER_H
#define DAYREPORTPARSER_H

#include "BaseReportParser.h"

/**
 * @brief 日报解析器
 * 处理格式：#Date:2024-01-01, #t#00:00, #d#RTU号
 */
class DayReportParser : public BaseReportParser
{
    Q_OBJECT

public:
    explicit DayReportParser(ReportDataModel* model, QObject* parent = nullptr);
    ~DayReportParser() override = default;

    // ===== 实现基类接口 =====
    bool scanAndParse() override;
    bool executeQueries(QProgressDialog* progress) override;
    void restoreToTemplate() override;
    void runCorrectnessTest() override;

    // ===== 日报特有接口 =====
    QString getBaseDate() const { return m_baseDate; }

protected:
    // ===== 实现纯虚函数 =====
    bool findDateMarker() override;
    void parseRow(int row) override;
    QTime getTaskTime(const QueryTask& task) override;
    QDateTime constructDateTime(const QString& date, const QString& time) override;
    int getQueryIntervalSeconds() const override { return 60; }  // 日报间隔60秒

    bool runAsyncTask() override;

private:
    // ===== 日报特有的标记识别 =====
    bool isDateMarker(const QString& text) const;
    QString extractDate(const QString& text);
};

#endif // DAYREPORTPARSER_H