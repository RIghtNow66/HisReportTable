#include "UnifiedQueryParser.h"
#include "reportdatamodel.h"
#include "TaosDataFetcher.h"
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include <QMessageBox>
#include <QFileInfo>
#include <QProgressDialog>
#include <cmath>

UnifiedQueryParser::UnifiedQueryParser(ReportDataModel* model, QObject* parent)
    : BaseReportParser(model, parent)
    , m_intervalSeconds(3600)
{
}

void UnifiedQueryParser::setTimeRange(const QDateTime& start, const QDateTime& end, int interval)
{
    m_startTime = start;
    m_endTime = end;
    m_intervalSeconds = interval;

    qDebug() << QString("设置时间范围：%1 ~ %2, 间隔%3秒")
        .arg(start.toString()).arg(end.toString()).arg(interval);
}

bool UnifiedQueryParser::scanAndParse()
{
    qDebug() << "========== 开始解析统一查询配置 ==========";

    // 从模型获取配置文件路径（在 loadReportTemplate 中设置）
    // 这里需要 model 提供一个接口返回文件路径
    // 暂时先假设通过构造时传入

    return true;  // 配置加载在 loadConfigFile 中完成
}

bool UnifiedQueryParser::loadConfigFile(const QString& filePath)
{
    qDebug() << "加载配置文件：" << filePath;

    // 1. 打开Excel
    QXlsx::Document xlsx(filePath);
    if (!xlsx.load()) {
        QMessageBox::warning(nullptr, "错误", "无法打开配置文件");
        return false;
    }

    QXlsx::Worksheet* sheet = static_cast<QXlsx::Worksheet*>(xlsx.currentSheet());
    if (!sheet) {
        return false;
    }

    // 2. 提取报表名称
    QFileInfo fileInfo(filePath);
    QString baseName = fileInfo.baseName();
    m_config.reportName = baseName.startsWith("#REPO_") ? baseName.mid(6) : baseName;
    m_config.configFilePath = filePath;

    // 3. 读取配置
    m_config.columns.clear();
    QXlsx::CellRange range = sheet->dimension();

    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        auto cellA = sheet->cellAt(row, 1);
        auto cellB = sheet->cellAt(row, 2);

        if (!cellA || !cellB) continue;

        QString columnName = cellA->value().toString().trimmed();
        QString rtuId = cellB->value().toString().trimmed();

        if (columnName.isEmpty() || rtuId.isEmpty()) continue;

        ReportColumnConfig colConfig;
        colConfig.displayName = columnName;
        colConfig.rtuId = rtuId;
        colConfig.sourceRow = row - 1;

        m_config.columns.append(colConfig);
    }

    qDebug() << QString("配置加载完成：%1 个数据列").arg(m_config.columns.size());
    return !m_config.columns.isEmpty();
}

bool UnifiedQueryParser::executeQueries(QProgressDialog* progress)
{
    if (m_config.columns.isEmpty()) {
        QMessageBox::warning(nullptr, "错误", "配置未加载");
        return false;
    }

    if (!m_startTime.isValid() || !m_endTime.isValid()) {
        QMessageBox::warning(nullptr, "错误", "时间范围未设置");
        return false;
    }

    // 1. 生成时间轴
    m_timeAxis = generateTimeAxis();

    if (progress) {
        progress->setLabelText("正在查询数据...");
        progress->setRange(0, 100);
        progress->setValue(10);
    }

    // 2. 构造查询地址
    QString queryAddr = buildQueryAddress();
    qDebug() << "查询地址：" << queryAddr;

    if (progress) progress->setValue(20);

    // 3. 执行查询
    std::map<int64_t, std::vector<float>> rawData;
    try {
        rawData = m_fetcher->fetchDataFromAddress(queryAddr.toStdString());
    }
    catch (const std::exception& e) {
        QMessageBox::critical(nullptr, "查询失败", QString("数据库错误：%1").arg(e.what()));
        return false;
    }

    if (progress) progress->setValue(60);

    if (rawData.empty()) {
        QMessageBox::warning(nullptr, "无数据", "查询未返回任何数据");
        return false;
    }

    // 4. 数据对齐
    m_alignedData = alignData(rawData);

    if (progress) progress->setValue(100);

    qDebug() << "查询完成";
    return true;
}

void UnifiedQueryParser::restoreToTemplate()
{
    m_timeAxis.clear();
    m_alignedData.clear();
    qDebug() << "统一查询数据已清空";
}

QString UnifiedQueryParser::buildQueryAddress()
{
    QStringList rtuList;
    for (const auto& col : m_config.columns) {
        rtuList.append(col.rtuId);
    }

    QString rtuPart = rtuList.join(",");
    QString startStr = m_startTime.toString("yyyy-MM-dd HH:mm:ss");
    QString endStr = m_endTime.toString("yyyy-MM-dd HH:mm:ss");

    return QString("%1@%2~%3#%4")
        .arg(rtuPart)
        .arg(startStr)
        .arg(endStr)
        .arg(m_intervalSeconds);
}

QVector<QDateTime> UnifiedQueryParser::generateTimeAxis()
{
    QVector<QDateTime> result;
    QDateTime current = m_startTime;

    while (current <= m_endTime) {
        result.append(current);
        current = current.addSecs(m_intervalSeconds);
    }

    qDebug() << QString("生成时间轴：%1 个点").arg(result.size());
    return result;
}

QHash<QString, QVector<double>> UnifiedQueryParser::alignData(
    const std::map<int64_t, std::vector<float>>& rawData)
{
    QHash<QString, QVector<double>> result;
    QStringList rtuList;

    for (const auto& col : m_config.columns) {
        rtuList.append(col.rtuId);
        result[col.rtuId] = QVector<double>(m_timeAxis.size(), NAN);
    }

    int64_t tolerance = (int64_t)(m_intervalSeconds * 1000) / 2;
    int matchCount = 0;

    for (int i = 0; i < m_timeAxis.size(); ++i) {
        int64_t targetTime = m_timeAxis[i].toMSecsSinceEpoch();

        int64_t closestTime = -1;
        int64_t minDiff = LLONG_MAX;

        for (auto it = rawData.begin(); it != rawData.end(); ++it) {
            int64_t diff = qAbs(it->first - targetTime);
            if (diff < minDiff) {
                minDiff = diff;
                closestTime = it->first;
            }
        }

        if (closestTime != -1 && minDiff <= tolerance) {
            const std::vector<float>& values = rawData.at(closestTime);

            for (int j = 0; j < rtuList.size() && j < (int)values.size(); ++j) {
                result[rtuList[j]][i] = values[j];
            }
            matchCount++;
        }
    }

    qDebug() << QString("数据对齐完成：匹配 %1/%2").arg(matchCount).arg(m_timeAxis.size());
    return result;
}