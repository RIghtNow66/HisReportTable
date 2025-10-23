#include "DayReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"

#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept>
#include <QProgressDialog>
#include <QtConcurrent>

DayReportParser::DayReportParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
{
}

bool DayReportParser::scanAndParse()
{
    qDebug() << "========== 开始解析日报 ==========";

    // 清空旧数据
    m_queryTasks.clear();
    m_dateFound = false;
    m_baseDate.clear();
    m_currentTime.clear();
    m_dataCache.clear();

    // 查找 #Date 标记
    if (!findDateMarker()) {
        QString errMsg = "错误：未找到 #Date 标记";
        qWarning() << errMsg;
        emit parseCompleted(false, errMsg);
        return false;
    }

    qDebug() << "基准日期:" << m_baseDate;

    // 逐行解析
    int totalRows = m_model->rowCount();
    for (int row = 0; row < totalRows; ++row) {
        parseRow(row);
        emit parseProgress(row + 1, totalRows);
    }

    if (m_queryTasks.isEmpty()) {
        QString warnMsg = "警告：未找到任何数据标记";
        qWarning() << warnMsg;
        emit parseCompleted(false, warnMsg);
        return false;
    }

    qDebug() << "解析完成：找到" << m_queryTasks.size() << "个数据点";

    setEditState(PREFETCHING);
    m_model->updateEditability();  // 通知 Model 更新可编辑状态

    // 启动后台预查询
    qDebug() << "========== 开始后台预查询 ==========";
    startAsyncTask();

    QString msg = QString("解析成功：找到 %1 个数据点，数据加载中...")
        .arg(m_queryTasks.size());
    emit parseCompleted(true, msg);

    qDebug() << "========================================";
    return true;
}

bool DayReportParser::runAsyncTask()
{
    qDebug() << "[后台线程] 日报预查询开始...";
    // analyzeAndPrefetch 包含了取消检查
    return this->analyzeAndPrefetch();
}

bool DayReportParser::findDateMarker()
{
    int totalRows = m_model->rowCount();
    int totalCols = m_model->columnCount();

    for (int row = 0; row < totalRows; ++row) {
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (!cell) continue;

            // ===== 使用 scanText() 而不是 value.toString() =====
            QString text = cell->scanText().trimmed();

            if (isDateMarker(text)) {
                m_baseDate = extractDate(text);

                if (m_baseDate.isEmpty()) {
                    return false;
                }

                m_dateFound = true;
                cell->cellType = CellData::DateMarker;
                cell->markerText = text;  // 保存原始标记

                // 设置显示格式
                QDate date = QDate::fromString(m_baseDate, "yyyy-MM-dd");
                if (date.isValid()) {
                    cell->displayValue = QString("%1年%2月%3日")
                        .arg(date.year())
                        .arg(date.month())
                        .arg(date.day());
                }
                else {
                    cell->displayValue = text;  // 无效日期保持原样
                }

                // 兼容性
                cell->originalMarker = text;

                qDebug() << QString("找到 #Date 标记: 行%1 列%2 → %3")
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

    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        // ===== 使用 scanText() =====
        QString text = cell->scanText().trimmed();
        if (text.isEmpty()) continue;

        // 遇到 #t# 时间标记
        if (isTimeMarker(text)) {
            QString timeStr = extractTime(text);
            m_currentTime = timeStr;

            cell->cellType = CellData::TimeMarker;
            cell->markerText = text;                    // 保存原始标记
            cell->displayValue = timeStr.left(5);       // 显示 "00:00"
            cell->originalMarker = text;                // 兼容性

            qDebug() << QString("行%1 列%2: 时间标记 %3 → %4")
                .arg(row).arg(col).arg(text).arg(m_currentTime);
        }
        // 遇到 #d# 数据标记
        else if (isDataMarker(text)) {
            if (m_currentTime.isEmpty()) {
                qWarning() << QString("行%1列%2 缺少时间信息，跳过").arg(row).arg(col);
                continue;
            }

            QString rtuId = extractRtuId(text);
            if (rtuId.isEmpty()) {
                qWarning() << QString("行%1列%2 RTU号为空，跳过").arg(row).arg(col);
                continue;
            }

            cell->cellType = CellData::DataMarker;
            cell->markerText = text;                    // 保存原始标记
            cell->rtuId = rtuId;
            cell->displayValue = text;                  // 初始显示标记
            cell->originalMarker = text;                // 兼容性

            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = "";

            m_queryTasks.append(task);

            qDebug() << QString("  行%1 列%2: RTU=%3, 时间=%4")
                .arg(row).arg(col).arg(rtuId).arg(m_currentTime);
        }
    }
}

QTime DayReportParser::getTaskTime(const QueryTask& task)
{
    int row = task.row;
    int totalCols = m_model->columnCount();

    // 在同一行查找时间标记
    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (cell && cell->cellType == CellData::TimeMarker) {
            QString timeStr = cell->displayValue.toString();
            QTime time = QTime::fromString(timeStr, "HH:mm:ss");
            if (!time.isValid()) {
                time = QTime::fromString(timeStr, "HH:mm");
            }
            return time;
        }
    }

    return QTime();
}

QDateTime DayReportParser::constructDateTime(const QString& date, const QString& time)
{
    QString dateTimeStr = date + " " + time;
    QDateTime result = QDateTime::fromString(dateTimeStr, "yyyy-MM-dd HH:mm:ss");

    if (!result.isValid()) {
        qWarning() << "日期时间构造失败:" << dateTimeStr;
    }

    return result;
}

bool DayReportParser::executeQueries(QProgressDialog* progress)
{

    if (m_queryTasks.isEmpty()) {
        return true;
    }

    bool cacheReady = analyzeAndPrefetch();

    if (!cacheReady) {
        qWarning() << "查询失败";
        // 不直接返回，继续用现有缓存填充
    }

    qDebug() << "========== 开始填充数据 ==========";


    if (progress) {
        progress->setRange(0, m_queryTasks.size());
        progress->setLabelText("正在填充数据...");
    }

    int successCount = 0;
    int failCount = 0;

    for (int i = 0; i < m_queryTasks.size(); ++i) {
        if (progress && progress->wasCanceled()) {
            for (int j = i; j < m_queryTasks.size(); ++j) {
                m_queryTasks[j].cell->queryExecuted = false;
            }
            return false;
        }

        const QueryTask& task = m_queryTasks[i];

        QTime time = getTaskTime(task);
        if (!time.isValid()) {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
            continue;
        }

        QDateTime dateTime = QDateTime(
            QDate::fromString(m_baseDate, "yyyy-MM-dd"),
            time
        );
        int64_t timestamp = dateTime.toMSecsSinceEpoch();

        float value = 0.0f;
        if (cacheReady && findInCache(task.cell->rtuId, timestamp, value)) {
            task.cell->value = QString::number(value, 'f', 2);
            task.cell->queryExecuted = true;
            task.cell->querySuccess = true;
            successCount++;
        }
        else {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
        }

        if (progress) {
            progress->setValue(i + 1);
        }
        emit queryProgress(i + 1, m_queryTasks.size());
    }

    qDebug() << QString("填充完成: 成功 %1, 失败 %2").arg(successCount).arg(failCount);
    emit queryCompleted(successCount, failCount);
    m_model->notifyDataChanged();

    return successCount > 0;
}

void DayReportParser::restoreToTemplate()
{
    qDebug() << "恢复到模板初始状态...";

    int restoredCount = 0;
    int nullCount = 0;

    for (const auto& task : m_queryTasks) {
        if (!task.cell) {
            nullCount++;
            qWarning() << QString("跳过无效单元格：行%1 列%2").arg(task.row).arg(task.col);
            continue;
        }

        // 还原为标记文本
        task.cell->displayValue = task.cell->markerText;  // 显示原始标记
        task.cell->queryExecuted = false;
        task.cell->querySuccess = false;
        restoredCount++;
    }

    qDebug() << QString("还原完成：成功%1个，跳过%2个无效单元格").arg(restoredCount).arg(nullCount);
}

bool DayReportParser::isDateMarker(const QString& text) const
{
    return text.startsWith("#Date:", Qt::CaseInsensitive);
}

QString DayReportParser::extractDate(const QString& text)
{
    QString dateStr = text.mid(6).trimmed();

    QDate date = QDate::fromString(dateStr, "yyyy-MM-dd");
    if (!date.isValid()) {
        qWarning() << "日期格式错误，期望 yyyy-MM-dd，实际:" << dateStr;
        return QString();
    }

    return dateStr;
}

QString DayReportParser::findTimeForDataMarker(int row, int col)
{
    // 在同一行查找时间标记
    int totalCols = m_model->columnCount();
    for (int c = 0; c < totalCols; ++c) {
        CellData* cell = m_model->getCell(row, c);
        if (cell && cell->cellType == CellData::TimeMarker) {
            return cell->displayValue.toString();
        }
    }
    return QString();
}