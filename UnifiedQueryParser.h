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
    ~UnifiedQueryParser();

    // ===== 实现基类接口 =====
    bool scanAndParse() override;
    bool executeQueries(QProgressDialog* progress) override;
    void restoreToTemplate() override;

    // ===== 统一查询特有接口 =====
    void setTimeRange(const TimeRangeConfig& config);
    const HistoryReportConfig& getConfig() const { return m_config; }
    const QVector<QDateTime>& getTimeAxis() const { return m_timeAxis; }
    const QHash<QString, QVector<double>>& getAlignedData() const { return m_alignedData; }

    int getQueryIntervalSeconds() const override { return m_timeConfig.intervalSeconds; }

    void startAsyncQuery();
    void requestCancel();
    bool isQuerying() const { return m_isQuerying; }

signals:
    // 查询进度信号（在后台线程中发射，主线程接收）
    void queryProgressUpdated(int current, int total);
    void queryStageChanged(QString stage);  // 查询阶段变化（如"正在生成时间轴"、"正在查询数据库"等）

    void queryCompletedWithStatus(bool success, QString message);

private slots:
    void onQueryFinished();

protected:
    // ===== 实现纯虚函数（统一查询不需要这些）=====
    bool findDateMarker() override { return true; }
    void parseRow(int row) override { Q_UNUSED(row); }
    QTime getTaskTime(const QueryTask& task) override { Q_UNUSED(task); return QTime(); }
    QDateTime constructDateTime(const QString& date, const QString& time) override {
        Q_UNUSED(date); Q_UNUSED(time); return QDateTime();
    }


    bool analyzeAndPrefetch() override { return true; }  // 空实现
    QList<TimeBlock> identifyTimeBlocks() override { return QList<TimeBlock>(); }  // 空实现

    bool executeQueriesInternal();

private:
    HistoryReportConfig m_config;           // 配置信息
    TimeRangeConfig m_timeConfig;           // 时间配置
    QVector<QDateTime> m_timeAxis;          // 时间轴
    QHash<QString, QVector<double>> m_alignedData;  // 对齐后的数据

    QFuture<bool> m_queryFuture;
    QFutureWatcher<bool>* m_queryWatcher;
    QAtomicInt m_cancelRequested;
    bool m_isQuerying;

private:
    // ===== 私有辅助函数 =====
    bool loadConfigFromCells();
    QString buildQueryAddress();
    QVector<QDateTime> generateTimeAxis();
    QHash<QString, QVector<double>> alignData(
        const std::map<int64_t, std::vector<float>>& rawData);
};

#endif // UNIFIEDQUERYPARSER_H