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
 * @brief �ձ�������
 *
 * ������� ##Day_xxx.xlsx ��ʽ���ձ��ļ�
 * ʶ�� #Date��#t#��#d# ��ǣ�������ѯ��ַ��ִ�в�ѯ
 */
class DayReportParser : public QObject
{
    Q_OBJECT

public:
    explicit DayReportParser(ReportDataModel* model, QObject* parent = nullptr);
    ~DayReportParser();

    // ===== ���Ľӿ� =====

    /**
     * @brief ɨ�貢�������е�Ԫ����
     * @return �Ƿ�ɹ��������ҵ�#Date���д���ѯ��Ԫ��
     */
    bool scanAndParse();

    /**
     * @brief ִ�����в�ѯ����
     * @return �Ƿ�ȫ���ɹ�
     */
    bool executeQueries();

    // ===== ״̬��ѯ =====

    bool isValid() const { return m_dateFound; }
    QString getBaseDate() const { return m_baseDate; }
    int getPendingQueryCount() const { return m_queryTasks.size(); }

signals:
    void parseProgress(int current, int total);      // ��������
    void queryProgress(int current, int total);      // ��ѯ����
    void parseCompleted(bool success, QString message);
    void queryCompleted(int successCount, int failCount);

private:
    // ===== ��һ�׶Σ���ǽ��� =====

    /**
     * @brief ���Ҳ����� #Date ���
     * @return �Ƿ��ҵ���Ч��#Date
     */
    bool findDateMarker();

    /**
     * @brief ����һ���е����б��
     * @param row �к�
     */
    void parseRow(int row);

    // ===== �ڶ��׶Σ���ѯִ�� =====

    /**
     * @brief ִ�е�����ѯ����
     * @param queryPath ��ѯ��ַ
     * @return ��ѯ�������ֵ��
     */
    QVariant querySinglePoint(const QString& queryPath);

    // ===== ���ߺ��� =====

    bool isDateMarker(const QString& text) const;
    bool isTimeMarker(const QString& text) const;
    bool isDataMarker(const QString& text) const;

    QString extractDate(const QString& text);        // �� "#Date:2024-01-01" ��ȡ
    QString extractTime(const QString& text);        // �� "#t#0:00" ��ȡ����ʽ��
    QString extractRtuId(const QString& text);       // �� "#d#AIRTU..." ��ȡ

    /**
     * @brief �����ѯ��ַ
     * @param rtuId RTU��
     * @param baseTime ��ʼʱ��
     * @return ��ѯ��ַ�ַ���
     */
    QString buildQueryPath(const QString& rtuId, const QDateTime& baseTime);

    /**
     * @brief ƴ�����ں�ʱ��
     * @param date "2024-01-01"
     * @param time "00:00:00"
     * @return QDateTime����
     */
    QDateTime constructDateTime(const QString& date, const QString& time);

private:
    // ===== ���ݳ�Ա =====

    ReportDataModel* m_model;              // ����ģ��ָ��
    TaosDataFetcher* m_fetcher;            // ���ݲ�ѯ��

    // ����״̬
    QString m_baseDate;                    // ��׼���ڣ��� #Date ��ȡ���� "2024-01-01"��
    QString m_currentTime;                 // ��ǰ��Чʱ�䣨������� #t# ��ȡ���� "00:00:00"��
    bool m_dateFound;                      // �Ƿ��ҵ� #Date ���

    // ��ѯ����
    struct QueryTask {
        CellData* cell;                    // ��Ԫ��ָ��
        int row;                           // �к�
        int col;                           // �к�
        QString queryPath;                 // ��ѯ��ַ
    };

    QList<QueryTask> m_queryTasks;         // ��ִ�еĲ�ѯ�����б�
};

#endif // DAYREPORTPARSER_H
