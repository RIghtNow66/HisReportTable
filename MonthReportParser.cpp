#include "MonthReportParser.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include "TaosDataFetcher.h"

#include <qDebug>
#include <QDate>
#include <QTime>
#include <stdexcept>
#include <QProgressDialog>
#include <QtConcurrent>
#include <QVariant>

MonthReportParser::MonthReportParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
{
}

bool MonthReportParser::scanAndParse()
{
    qDebug() << "========== 开始解析月报 ==========";

    // 清空旧数据
    m_queryTasks.clear();
    m_dateFound = false;
    m_baseYearMonth.clear();
    m_baseTime.clear();
    m_currentTime.clear();
    m_dataCache.clear();

    // 查找 #Date1 和 #Date2 标记
    if (!findDateMarker()) {
        QString errMsg = "错误：未找到 #Date1 或 #Date2 标记";
        qWarning() << errMsg;
        emit parseCompleted(false, errMsg);
        return false;
    }

    qDebug() << "基准年月:" << m_baseYearMonth;
    qDebug() << "基准时间:" << m_baseTime;

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

    collectActualDays();

    // 启动后台预查询
    qDebug() << "========== 开始后台预查询 ==========";
    m_isPrefetching = true;

    m_prefetchFuture = QtConcurrent::run([this]() -> bool {
        qDebug() << "[后台线程] 月报预查询开始...";
        try {
            bool result = this->analyzeAndPrefetch();
            qDebug() << "[后台线程] 月报预查询" << (result ? "成功" : "失败");
            return result;
        }
        catch (const std::exception& e) {
            qWarning() << "[后台线程] 异常：" << e.what();
            return false;
        }
        });

    m_prefetchWatcher->setFuture(m_prefetchFuture);

    QString msg = QString("解析成功：找到 %1 个数据点，数据加载中...")
        .arg(m_queryTasks.size());
    emit parseCompleted(true, msg);

    qDebug() << "========================================";
    return true;
}

bool MonthReportParser::findDateMarker()
{
    int totalRows = m_model->rowCount();
    int totalCols = m_model->columnCount();

    bool foundDate1 = false;
    bool foundDate2 = false;

    for (int row = 0; row < totalRows; ++row) {
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (!cell) continue;

            QString text = cell->value.toString().trimmed();

            // 查找 #Date1:2024-01
            if (isDate1Marker(text)) {
                m_baseYearMonth = extractYearMonth(text);

                if (m_baseYearMonth.isEmpty()) {
                    qWarning() << "年月格式错误:" << text;
                    return false;
                }

                cell->cellType = CellData::DateMarker;
                cell->originalMarker = text;

                // 设置显示格式：2024年1月
                QDate date = QDate::fromString(m_baseYearMonth + "-01", "yyyy-MM-dd");
                if (date.isValid()) {
                    cell->value = QString("%1年%2月")
                        .arg(date.year())
                        .arg(date.month());
                }

                qDebug() << QString("找到 #Date1 标记: 行%1 列%2 → %3")
                    .arg(row).arg(col).arg(m_baseYearMonth);

                foundDate1 = true;
            }
            // 查找 #Date2:08:30
            else if (isDate2Marker(text)) {
                m_baseTime = extractTimeOfDay(text);

                if (m_baseTime.isEmpty()) {
                    qWarning() << "时间格式错误:" << text;
                    return false;
                }

                cell->cellType = CellData::TimeMarker;
                cell->originalMarker = text;
                cell->value = m_baseTime.left(5);  // "08:30:00" → "08:30"

                qDebug() << QString("找到 #Date2 标记: 行%1 列%2 → %3")
                    .arg(row).arg(col).arg(m_baseTime);

                foundDate2 = true;
            }

            // 如果两个都找到了，可以提前退出
            if (foundDate1 && foundDate2) {
                m_dateFound = true;
                // 将年月和时间组合存入 m_baseDate，供基类使用
                m_baseDate = m_baseYearMonth;  // 基类的查询需要这个
                return true;
            }
        }
    }

    if (!foundDate1) {
        qWarning() << "未找到 #Date1 标记";
    }
    if (!foundDate2) {
        qWarning() << "未找到 #Date2 标记";
    }

    return foundDate1 && foundDate2;
}

void MonthReportParser::parseRow(int row)
{
    int totalCols = m_model->columnCount();

    for (int col = 0; col < totalCols; ++col) {
        CellData* cell = m_model->getCell(row, col);
        if (!cell) continue;

        QString text = cell->value.toString().trimmed();
        if (text.isEmpty()) continue;

        // 遇到 #t# 日期标记（月报中表示"日"）
        if (isTimeMarker(text)) {
            int day = extractDay(text);

            // 验证日期是否有效
            QDate date = QDate::fromString(m_baseYearMonth + QString("-%1").arg(day, 2, 10, QChar('0')), "yyyy-MM-dd");
            if (!date.isValid()) {
                // 无效日期（如2月30日），保留原始标记，后续查询会返回N/A
                qWarning() << QString("行%1列%2: 无效日期 %3-%4")
                    .arg(row).arg(col).arg(m_baseYearMonth).arg(day);
                cell->cellType = CellData::TimeMarker;
                cell->originalMarker = text;
                cell->value = text;  // 显示原始文本
                continue;
            }

            // 有效日期，存储日期信息
            m_currentTime = QString("%1").arg(day);  // 存储日期数字

            cell->cellType = CellData::TimeMarker;
            cell->originalMarker = text;
            cell->value = QString::number(day);  // 显示日期数字

            qDebug() << QString("行%1 列%2: 日期标记 #t#%3 → %4日")
                .arg(row).arg(col).arg(day).arg(day);
        }
        // 遇到 #d# 数据标记
        else if (isDataMarker(text)) {
            if (m_currentTime.isEmpty()) {
                qWarning() << QString("行%1列%2 缺少日期信息，跳过").arg(row).arg(col);
                continue;
            }

            QString rtuId = extractRtuId(text);
            if (rtuId.isEmpty()) {
                qWarning() << QString("行%1列%2 RTU号为空，跳过").arg(row).arg(col);
                continue;
            }

            cell->cellType = CellData::DataMarker;
            cell->originalMarker = text;
            cell->rtuId = rtuId;

            QueryTask task;
            task.cell = cell;
            task.row = row;
            task.col = col;
            task.queryPath = "";

            m_queryTasks.append(task);

            qDebug() << QString("  行%1 列%2: RTU=%3, 日期=%4日 %5")
                .arg(row).arg(col).arg(rtuId).arg(m_currentTime).arg(m_baseTime);
        }
    }
}

QTime MonthReportParser::getTaskTime(const QueryTask& task)
{
    // 月报中，时间信息固定为 m_baseTime
    // 这里只需要返回基准时间即可
    QTime time = QTime::fromString(m_baseTime, "HH:mm:ss");
    if (!time.isValid()) {
        time = QTime::fromString(m_baseTime, "HH:mm");
    }
    return time;
}

QDateTime MonthReportParser::constructDateTime(const QString& date, const QString& time)
{
    // date: "2024-01-15" (年月日)
    // time: "08:30:00" (时分秒)
    QString dateTimeStr = date + " " + time;
    QDateTime result = QDateTime::fromString(dateTimeStr, "yyyy-MM-dd HH:mm:ss");

    if (!result.isValid()) {
        qWarning() << "月报日期时间构造失败:" << dateTimeStr;
    }

    return result;
}

bool MonthReportParser::executeQueries(QProgressDialog* progress)
{
    if (m_queryTasks.isEmpty()) {
        return true;
    }

    qDebug() << "========== 开始填充月报数据 ==========";

    bool cacheReady = true;
    if (m_dataCache.isEmpty()) {
        qWarning() << "缓存为空，重新查询...";
        cacheReady = analyzeAndPrefetch();

        if (!cacheReady) {
            qWarning() << "查询失败，将所有数据标记为N/A";
        }
    }

    if (progress) {
        progress->setRange(0, m_queryTasks.size());
        progress->setLabelText("正在填充月报数据...");
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

        // 从任务所在行找到日期标记
        int day = 0;
        int totalCols = m_model->columnCount();
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(task.row, col);
            if (cell && cell->cellType == CellData::TimeMarker) {
                QString dayStr = cell->value.toString();
                day = dayStr.toInt();
                break;
            }
        }

        if (day == 0) {
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
            continue;
        }

        // 构造完整日期
        QString fullDate = QString("%1-%2").arg(m_baseYearMonth).arg(day, 2, 10, QChar('0'));

        // 验证日期有效性
        QDate date = QDate::fromString(fullDate, "yyyy-MM-dd");
        if (!date.isValid()) {
            // 无效日期（如2月30日）
            task.cell->value = "N/A";
            task.cell->queryExecuted = true;
            task.cell->querySuccess = false;
            failCount++;
            continue;
        }

        // 构造时间戳
        QDateTime dateTime = constructDateTime(fullDate, m_baseTime);
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

    qDebug() << QString("月报填充完成: 成功 %1, 失败 %2").arg(successCount).arg(failCount);
    emit queryCompleted(successCount, failCount);
    m_model->notifyDataChanged();

    return successCount > 0;
}

void MonthReportParser::restoreToTemplate()
{
    qDebug() << "恢复月报到模板初始状态...";
    m_dataCache.clear();
    for (const auto& task : m_queryTasks) {
        if (task.cell) {
            task.cell->value = task.cell->originalMarker;
            task.cell->queryExecuted = false;
            task.cell->querySuccess = false;
        }
    }
}

QString MonthReportParser::extractTime(const QString& text)
{
    // 月报中 #t# 后面是日期数字，不是时间
    // 例如：#t#1 表示1日
    QString dayStr = text.mid(3).trimmed();
    return dayStr;
}

// ===== 私有辅助函数 =====

bool MonthReportParser::isDate1Marker(const QString& text) const
{
    return text.startsWith("#Date1:", Qt::CaseInsensitive);
}

bool MonthReportParser::isDate2Marker(const QString& text) const
{
    return text.startsWith("#Date2:", Qt::CaseInsensitive);
}

QString MonthReportParser::extractYearMonth(const QString& text)
{
    // #Date1:2024-01 → "2024-01"
    QString yearMonth = text.mid(7).trimmed();

    // 验证格式
    QRegularExpression regex("^\\d{4}-\\d{2}$");
    if (!regex.match(yearMonth).hasMatch()) {
        qWarning() << "年月格式错误，期望 yyyy-MM，实际:" << yearMonth;
        return QString();
    }

    // 验证是否为有效月份
    QDate testDate = QDate::fromString(yearMonth + "-01", "yyyy-MM-dd");
    if (!testDate.isValid()) {
        qWarning() << "无效的年月:" << yearMonth;
        return QString();
    }

    return yearMonth;
}

QString MonthReportParser::extractTimeOfDay(const QString& text)
{
    // #Date2:08:30 → "08:30:00"
    QString timeStr = text.mid(7).trimmed();

    QStringList parts = timeStr.split(":");
    if (parts.size() == 2) {
        // "08:30" → "08:30:00"
        timeStr += ":00";
    }
    else if (parts.size() != 3) {
        qWarning() << "时间格式错误，期望 HH:mm 或 HH:mm:ss，实际:" << timeStr;
        return QString();
    }

    // 验证时间有效性
    QTime time = QTime::fromString(timeStr, "HH:mm:ss");
    if (!time.isValid()) {
        // 尝试宽松格式
        time = QTime::fromString(timeStr, "H:mm:ss");
        if (!time.isValid()) {
            qWarning() << "时间解析失败:" << timeStr;
            return QString();
        }
    }

    return time.toString("HH:mm:ss");
}

int MonthReportParser::extractDay(const QString& text)
{
    // #t#1 → 1
    // #t#15 → 15
    QString dayStr = text.mid(3).trimmed();

    bool ok;
    int day = dayStr.toInt(&ok);

    if (!ok || day < 1 || day > 31) {
        qWarning() << "日期数字格式错误:" << text;
        return 0;
    }

    return day;
}

// 收集实际日期
void MonthReportParser::collectActualDays()
{
    m_actualDays.clear();

    qDebug() << "开始收集月报实际日期...";

    for (int row = 0; row < m_model->rowCount(); ++row) {
        for (int col = 0; col < m_model->columnCount(); ++col) {
            CellData* cell = m_model->getCell(row, col);
            if (cell && cell->cellType == CellData::TimeMarker) {
                QString dayStr = cell->value.toString();
                bool ok;
                int day = dayStr.toInt(&ok);

                if (ok && day >= 1 && day <= 31) {
                    QString fullDate = QString("%1-%2").arg(m_baseYearMonth).arg(day, 2, 10, QChar('0'));
                    QDate date = QDate::fromString(fullDate, "yyyy-MM-dd");

                    if (date.isValid()) {
                        m_actualDays.insert(day);
                    }
                    else {
                        qWarning() << QString("跳过无效日期：%1").arg(fullDate);
                    }
                }
            }
        }
    }

    QList<int> sortedDays = m_actualDays.values();
    std::sort(sortedDays.begin(), sortedDays.end());

    QString daysStr;
    for (int i = 0; i < sortedDays.size(); ++i) {
        if (i > 0) daysStr += ", ";
        daysStr += QString::number(sortedDays[i]);
    }

    qDebug() << QString("收集到 %1 个有效日期：[%2]")
        .arg(m_actualDays.size())
        .arg(daysStr);
}


// 获取最小日期
int MonthReportParser::getMinDay() const
{
    if (m_actualDays.isEmpty()) return 1;
    return *std::min_element(m_actualDays.begin(), m_actualDays.end());
}

// 获取最大日期
int MonthReportParser::getMaxDay() const
{
    if (m_actualDays.isEmpty()) return 1;
    return *std::max_element(m_actualDays.begin(), m_actualDays.end());
}

// 识别时间块（月报版本）
QList<BaseReportParser::TimeBlock> MonthReportParser::identifyTimeBlocks()
{
    QList<TimeBlock> blocks;

    if (m_actualDays.isEmpty()) {
        qWarning() << "未找到任何有效日期标记";
        return blocks;
    }

    int minDay = getMinDay();
    int maxDay = getMaxDay();

    QString startDate = QString("%1-%2").arg(m_baseYearMonth).arg(minDay, 2, 10, QChar('0'));
    QString endDate = QString("%1-%2").arg(m_baseYearMonth).arg(maxDay, 2, 10, QChar('0'));

    qDebug() << QString("月报查询范围：%1 ~ %2").arg(startDate).arg(endDate);

    QTime baseTime = QTime::fromString(m_baseTime, "HH:mm:ss");
    QTime endTime = baseTime.addSecs(60);

    TimeBlock block;
    block.startTime = baseTime;
    block.endTime = endTime;
    block.startDate = startDate;
    block.endDate = endDate;

    m_currentQueryStartDate = startDate;
    m_currentQueryEndDate = endDate;

    blocks.append(block);

    qDebug() << QString("生成查询块：%1 %2 ~ %3 %4")
        .arg(startDate)
        .arg(baseTime.toString("HH:mm:ss"))
        .arg(endDate)
        .arg(endTime.toString("HH:mm:ss"));

    return blocks;
}


// 实现日期范围获取
bool MonthReportParser::getDateRange(QString& startDate, QString& endDate)
{
    if (m_currentQueryStartDate.isEmpty() || m_currentQueryEndDate.isEmpty()) {
        return false;
    }

    startDate = m_currentQueryStartDate;
    endDate = m_currentQueryEndDate;
    return true;
}

void MonthReportParser::runCorrectnessTest()
{
    qDebug() << "";
    qDebug() << "╔════════════════════════════════════════════════════════════╗";
    qDebug() << "║          月报数据对应关系完整性测试                        ║";
    qDebug() << "╚════════════════════════════════════════════════════════════╝";
    qDebug() << "";

    if (m_queryTasks.isEmpty()) {
        qDebug() << " 无测试数据";
        return;
    }

    // 选择前3个任务作为测试样本
    int testCount = qMin(3, m_queryTasks.size());
    QList<QueryTask> testTasks;
    for (int i = 0; i < testCount; ++i) {
        testTasks.append(m_queryTasks[i]);
    }

    qDebug() << "【测试样本】选择前" << testCount << "个数据点：";
    for (int i = 0; i < testTasks.size(); ++i) {
        const QueryTask& task = testTasks[i];

        // 找到对应的日期
        int day = 0;
        int totalCols = m_model->columnCount();
        for (int col = 0; col < totalCols; ++col) {
            CellData* cell = m_model->getCell(task.row, col);
            if (cell && cell->cellType == CellData::TimeMarker) {
                day = cell->value.toInt();
                break;
            }
        }

        qDebug() << QString("  样本%1: 行%2列%3, RTU=%4, 日期=%5-%6日 %7")
            .arg(i + 1)
            .arg(task.row + 1)
            .arg(task.col + 1)
            .arg(task.cell->rtuId)
            .arg(m_baseYearMonth)
            .arg(day)
            .arg(m_baseTime);
    }

    qDebug() << "";
    qDebug() << "【查询验证】";
    qDebug() << QString("  年月: %1").arg(m_baseYearMonth);
    qDebug() << QString("  时间: %1").arg(m_baseTime);
    qDebug() << QString("  查询间隔: %1 秒 (24小时)").arg(getQueryIntervalSeconds());
    qDebug() << QString("  预期返回: 该月天数 × RTU数量 个数据点");

    qDebug() << "";
    qDebug() << "╔════════════════════════════════════════════════════════════╗";
    qDebug() << "║   月报测试完成                                             ║";
    qDebug() << "╚════════════════════════════════════════════════════════════╝";
}