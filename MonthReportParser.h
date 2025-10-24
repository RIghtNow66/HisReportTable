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
    void restoreToTemplate() override {};

    // ===== 月报特有接口 =====
    QString getBaseYearMonth() const { return m_baseYearMonth; }
    QString getBaseTime() const { return m_baseTime; }

    // ===== 重写标记识别 =====
    QString extractTime(const QString& text) const override;

    QVariant formatDisplayValueForMarker(const CellData* cell) const override;

    void collectActualDays();
protected:
    // ===== 实现纯虚函数 =====
    bool findDateMarker() override;
    void parseRow(int row) override;
    QTime getTaskTime(const QueryTask& task) override;
    QDateTime constructDateTime(const QString& date, const QString& time) override;
    int getQueryIntervalSeconds() const override { return 60; }  // 月报间隔24小时

    QList<TimeBlock> identifyTimeBlocks() override;
    bool getDateRange(QString& startDate, QString& endDate) override;

    // ==== = 重写预查询逻辑 ==== =
    bool analyzeAndPrefetch() override;

    bool runAsyncTask() override;

    QString findTimeForDataMarker(int row, int col) override;

    void  onRescanCompleted(int newCount, int modifiedCount, int removedCount,
        const QSet<int>& affectedRows)  override;

private:
    // ===== 月报特有的标记识别 =====
    bool isDate1Marker(const QString& text) const;
    bool isDate2Marker(const QString& text) const;
    QString extractYearMonth(const QString& text) const;
    QString extractTimeOfDay(const QString& text) const;
    int extractDay(const QString& text) const;

    int getMinDay() const;
    int getMaxDay() const;

    /**
   * @brief 增量更新实际日期集合（仅处理变化的行）
   * @param affectedRows 受影响的行号集合
   */
    void updateActualDaysIncremental(const QSet<int>& affectedRows);

    /**
     * @brief 从指定行提取日期
     * @param row 行号
     * @return 日期数字（1-31），失败返回0
     */
    int extractDayFromRow(int row) const;

    /**
 * @brief 验证并清理 m_actualDays 中的无效日期
 */
    void validateActualDays();

private:
    QString m_baseYearMonth;   // "2024-01"
    QString m_baseTime;        // "08:30:00"

    QSet<int> m_actualDays; // 实际出现的日期集合：{ 10, 11, 12, ..., 20 }

    // 临时存储当前查询的日期范围（用于 getDateRange）
    QString m_currentQueryStartDate;
    QString m_currentQueryEndDate;
};

#endif // MONTHREPORTPARSER_H
