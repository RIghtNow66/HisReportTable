#pragma once
#ifndef UNIFIEDQUERYPARSER_H
#define UNIFIEDQUERYPARSER_H

#include "BaseReportParser.h"
#include "DataBindingConfig.h"

class UnifiedQueryParser : public BaseReportParser
{
    Q_OBJECT

public:
    explicit UnifiedQueryParser(ReportDataModel* model, QObject* parent = nullptr);
    ~UnifiedQueryParser() override = default;

    // ===== 实现基类接口 =====
    bool scanAndParse() override;
    bool executeQueries(QProgressDialog* progress) override;
    void restoreToTemplate() override;

    // ===== 统一查询特有接口 =====
    const HistoryReportConfig& getConfig() const { return m_config; }
    const QVector<QDateTime>& getTimeAxis() const { return m_timeAxis; }
    const QHash<QString, QVector<double>>& getAlignedData() const { return m_alignedData; }

    // 设置时间范围（从外部传入）
    void setTimeRange(const QDateTime& start, const QDateTime& end, int interval);

protected:
    // ===== 实现纯虚函数（但统一查询不需要这些）=====
    bool findDateMarker() override { return true; }  // 不需要
    void parseRow(int row) override { Q_UNUSED(row); }  // 不需要
    QTime getTaskTime(const QueryTask& task) override { Q_UNUSED(task); return QTime(); }
    QDateTime constructDateTime(const QString& date, const QString& time) override {
        Q_UNUSED(date); Q_UNUSED(time); return QDateTime();
    }
    int getQueryIntervalSeconds() const override { return m_intervalSeconds; }

private:
    // ===== 私有辅助函数 =====
    bool loadConfigFile(const QString& filePath);
    QString buildQueryAddress();
    QVector<QDateTime> generateTimeAxis();
    QHash<QString, QVector<double>> alignData(
        const std::map<int64_t, std::vector<float>>& rawData);

private:
    HistoryReportConfig m_config;           // 配置信息
    QVector<QDateTime> m_timeAxis;          // 时间轴
    QHash<QString, QVector<double>> m_alignedData;  // 对齐后的数据

    // 时间范围（从外部传入）
    QDateTime m_startTime;
    QDateTime m_endTime;
    int m_intervalSeconds;
};

#endif // UNIFIEDQUERYPARSER_H