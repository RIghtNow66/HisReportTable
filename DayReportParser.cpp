#include "DayReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"
#include <QDebug>
#include <QDate>
#include <QTime>

DayReportParser::DayReportParser(ReportDataModel* model, QObject* parent)
    : QObject(parent)
    , m_model(model)
    , m_fetcher(new TaosDataFetcher())
    , m_dateFound(false)
{
}

DayReportParser::~DayReportParser()
{
    delete m_fetcher;
}

// ===== �������� =====

bool DayReportParser::scanAndParse()
{
    qDebug() << "========== ��ʼ�����ձ� ==========";

    // ��վ�����
    m_queryTasks.clear();
    m_dateFound = false;
    m_baseDate.clear();
    m_currentTime.clear();

    // ��1�������� #Date ���
    if (!findDateMarker()) {
        QString errMsg = "����δ�ҵ� #Date ��ǣ��޷�ȷ����ѯ����";
        qWarning() << errMsg;
        emit parseCompleted(false, errMsg);
        return false;
    }

    qDebug() << "��׼����:" << m_baseDate;

    // ��2�������н������
    int totalRows = m_model->rowCount();

    for (int row = 0; row < totalRows; ++row) {
        parseRow(row);
        emit parseProgress(row + 1, totalRows);
    }

    // ��3���������
    if (m_queryTasks.isEmpty()) {
        QString warnMsg = "���棺δ�ҵ��κ� #d# ���ݱ��";
        qWarning() << warnMsg;
        emit parseCompleted(false, warnMsg);
        return false;
    }

    QString successMsg = QString("�����ɹ����ҵ� %1 ������ѯ��Ԫ��").arg(m_queryTasks.size());
    qDebug() << successMsg;
    qDebug() << "========================================";

    emit parseCompleted(true, successMsg);
    return true;
}

bool DayReportParser::findDateMarker()
{
    int totalRows = m_model->rowCount();
    int totalCols = m_model->columnCount();

    for (int row = 0; row < totalRows; ++row) {
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (!cell) continue;

            QString text = cell->value.toString().trimmed();

            if (isDateMarker(text)) {
                m_baseDate = extractDate(text);

                if (m_baseDate.isEmpty()) {
                    qWarning() << "���ڸ�ʽ����:" << text;
                    return false;
                }

                m_dateFound = true;

                // ��ǵ�Ԫ������
                cell->cellType = CellData::DateMarker;
                cell->originalMarker = text;

                qDebug() << QString("�ҵ� #Date ���: ��%1 ��%2 �� %3")
                    .arg(row).arg(col).arg(m_baseDate);

                return true;
            }
        }
    }

    return false;
}

void DayReportParser::parseRow(int row)
{
    int totalCols = m_model->columnCount();

    // ������ɨ��
    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->value.toString().trimmed();
        if (text.isEmpty()) continue;

        // ===== ���1������ #t# ʱ���� =====
        if (isTimeMarker(text)) {
            QString timeStr = extractTime(text);
            m_currentTime = timeStr;

            // ��ǵ�Ԫ������
            cell->cellType = CellData::TimeMarker;
            cell->originalMarker = text;

            // ������ʾֵ��ȥ�� #t# ǰ׺��ֻ��ʾ HH:mm��
            cell->value = timeStr.left(5);  // "00:00:00" �� "00:00"

            qDebug() << QString("  ��%1 ��%2: ʱ���� %3 �� %4")
                .arg(row).arg(col).arg(text).arg(m_currentTime);
        }
        // ===== ���2������ #d# ���ݱ�� =====
        else if (isDataMarker(text)) {
            // ���ǰ������
            if (m_currentTime.isEmpty()) {
                qWarning() << QString("���棺��%1��%2 ȱ��ʱ����Ϣ������").arg(row).arg(col);
                continue;
            }

            QString rtuId = extractRtuId(text);
            if (rtuId.isEmpty()) {
                qWarning() << QString("���棺��%1��%2 RTU��Ϊ�գ�����").arg(row).arg(col);
                continue;
            }

            // ��ǵ�Ԫ������
            cell->cellType = CellData::DataMarker;
            cell->originalMarker = text;
            cell->rtuId = rtuId;

            // ������ѯ��ַ
            QDateTime startTime = constructDateTime(m_baseDate, m_currentTime);
            QString queryPath = buildQueryPath(rtuId, startTime);

            cell->queryPath = queryPath;

            // ��ӵ���ѯ�����б�
            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = queryPath;
            m_queryTasks.append(task);

            qDebug() << QString("  ��%1 ��%2: ���ݱ�� %3 �� ��ѯ��ַ: %4")
                .arg(row).arg(col).arg(text).arg(queryPath);
        }
    }
}

bool DayReportParser::executeQueries()
{
    if (m_queryTasks.isEmpty()) {
        qWarning() << "û�д���ѯ������";
        return false;
    }

    qDebug() << "========== ��ʼִ�в�ѯ ==========";
    qDebug() << "����ѯ����:" << m_queryTasks.size();

    int successCount = 0;
    int failCount = 0;
    int total = m_queryTasks.size();

    for (int i = 0; i < total; ++i) {
        const QueryTask& task = m_queryTasks[i];

        qDebug() << QString("[%1/%2] ��ѯ: %3")
            .arg(i + 1).arg(total).arg(task.queryPath);

        try {
            QVariant result = querySinglePoint(task.queryPath);

            // ���µ�Ԫ��
            task.cell->value = result;
            task.cell->queryExecuted = true;
            task.cell->querySuccess = true;

            successCount++;

            qDebug() << QString(" �ɹ�: ��%1��%2 = %3")
                .arg(task.row).arg(task.col).arg(result.toString());
        }
        catch (const std::exception& e) {
            qWarning() << QString("ʧ��: ��%1��%2 - %3")
                .arg(task.row).arg(task.col).arg(e.what());

            // ��ѯʧ�ܣ���ʾ ERROR
            task.cell->value = "ERROR";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;

            failCount++;
        }

        emit queryProgress(i + 1, total);
    }

    qDebug() << "========================================";
    qDebug() << QString("��ѯ���: �ɹ� %1, ʧ�� %2").arg(successCount).arg(failCount);

    // ֪ͨģ��ˢ����ʾ
    m_model->notifyDataChanged();

    emit queryCompleted(successCount, failCount);

    return (failCount == 0);
}

QVariant DayReportParser::querySinglePoint(const QString& queryPath)
{
    // ���� TaosDataFetcher ִ�в�ѯ
    auto dataMap = m_fetcher->fetchDataFromAddress(queryPath.toStdString());

    if (dataMap.empty()) {
        throw std::runtime_error("��ѯ������");
    }

    // ȡ��һ��ʱ���ĵ�һ��ֵ
    auto firstEntry = dataMap.begin();
    if (firstEntry->second.empty()) {
        throw std::runtime_error("����Ϊ��");
    }

    float value = firstEntry->second[0];

    // ת��Ϊdouble������2λС��
    double result = static_cast<double>(value);
    return QString::number(result, 'f', 2);
}

// ===== ���ʶ�� =====

bool DayReportParser::isDateMarker(const QString& text) const
{
    return text.startsWith("#Date:", Qt::CaseInsensitive);
}

bool DayReportParser::isTimeMarker(const QString& text) const
{
    return text.startsWith("#t#", Qt::CaseInsensitive);
}

bool DayReportParser::isDataMarker(const QString& text) const
{
    return text.startsWith("#d#", Qt::CaseInsensitive);
}

// ===== ��Ϣ��ȡ =====

QString DayReportParser::extractDate(const QString& text)
{
    // "#Date:2024-01-01" �� "2024-01-01"
    QString dateStr = text.mid(6).trimmed();  // ���� "#Date:"

    // ��֤���ڸ�ʽ
    QDate date = QDate::fromString(dateStr, "yyyy-MM-dd");
    if (!date.isValid()) {
        qWarning() << "���ڸ�ʽ�������� yyyy-MM-dd��ʵ��:" << dateStr;
        return QString();
    }

    return dateStr;
}

QString DayReportParser::extractTime(const QString& text)
{
    // "#t#0:00" �� "00:00:00"
    // "#t#13:30" �� "13:30:00"

    QString timeStr = text.mid(3).trimmed();  // ���� "#t#"

    // ����ͬ��ʽ
    QStringList parts = timeStr.split(":");

    if (parts.size() == 2) {
        // "0:00" �� "00:00:00"
        timeStr += ":00";
    }
    else if (parts.size() != 3) {
        qWarning() << "ʱ���ʽ����:" << text;
        return "00:00:00";
    }

    // ���Խ�������ʽ��
    QTime time = QTime::fromString(timeStr, "H:mm:ss");
    if (!time.isValid()) {
        qWarning() << "ʱ�����ʧ��:" << timeStr;
        return "00:00:00";
    }

    // ���ر�׼��ʽ "HH:mm:ss"
    return time.toString("HH:mm:ss");
}

QString DayReportParser::extractRtuId(const QString& text)
{
    // "#d#AIRTU034700019" �� "AIRTU034700019"
    return text.mid(3).trimmed();
}

// ===== ��ѯ��ַ���� =====

QString DayReportParser::buildQueryPath(const QString& rtuId, const QDateTime& baseTime)
{
    // ��ʼʱ�䣺�û�ָ����ʱ��
    QString startTime = baseTime.toString("yyyy-MM-dd HH:mm:ss");

    // ����ʱ�䣺��ʼʱ�� + 1����
    QDateTime endTime = baseTime.addSecs(60);
    QString endTimeStr = endTime.toString("yyyy-MM-dd HH:mm:ss");

    // ��ѯ��ʽ��RTU��@��ʼʱ��~����ʱ��#���(��)
    // ʾ����AIRTU034700019@2024-01-01 00:00:00~2024-01-01 00:01:00#1
    return QString("%1@%2~%3#1")
        .arg(rtuId)
        .arg(startTime)
        .arg(endTimeStr);
}

QDateTime DayReportParser::constructDateTime(const QString& date, const QString& time)
{
    // "2024-01-01" + "00:00:00" �� QDateTime
    QString dateTimeStr = date + " " + time;
    QDateTime result = QDateTime::fromString(dateTimeStr, "yyyy-MM-dd HH:mm:ss");

    if (!result.isValid()) {
        qWarning() << "����ʱ�乹��ʧ��:" << dateTimeStr;
    }

    return result;
}