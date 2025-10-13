#include "reportdatamodel.h"
#include "formulaengine.h"
#include "excelhandler.h" // 用于文件操作
#include "UniversalQueryEngine.h"
// QXlsx相关（检查是否已包含）
#include "xlsxdocument.h"      // 用于 QXlsx::Document
#include "xlsxcellrange.h"     // 用于 QXlsx::CellRange
#include "xlsxworksheet.h"     // 用于 QXlsx::Worksheet
#include "xlsxformat.h"        // 用于 QXlsx::Format

#include <QColor>
#include <QFont>
#include <QPoint>
#include <QHash>
#include <QBrush>
#include <QPen>
#include <QFileInfo>           // 用于 QFileInfo
#include <QVariant>            // 用于 QVariant（可能已有）
#include <QProgressDialog>     // 用于 QProgressDialog
#include <QApplication>        // 用于 qApp
#include <QDebug>              // 用于 qDebug
#include <limits>              // 用于 std::numeric_limits
#include <cmath>               // 用于 std::isnan, std::isinf
#include <QMessageBox>


ReportDataModel::ReportDataModel(QObject* parent)
    : QAbstractTableModel(parent)
    , m_currentMode(REALTIME_MODE)
    , m_maxRow(100) // 默认初始行数
    , m_maxCol(26)  // 默认初始列数 (A-Z)
    , m_formulaEngine(new FormulaEngine(this))
{
}

ReportDataModel::~ReportDataModel()
{
    qDeleteAll(m_cells);
}

// --- Qt Model 核心接口实现 ---

int ReportDataModel::rowCount(const QModelIndex& parent) const
{
    Q_UNUSED(parent)
        return m_maxRow;
}

int ReportDataModel::columnCount(const QModelIndex& parent) const
{
    Q_UNUSED(parent)
        return m_maxCol;
}

QSize ReportDataModel::span(const QModelIndex& index) const
{
    if (!index.isValid())
        return QSize(1, 1);

    const CellData* cell = getCell(index.row(), index.column());
    if (!cell || !cell->mergedRange.isValid() || !cell->mergedRange.isMerged())
        return QSize(1, 1);

    // 只有主单元格返回span信息
    if (index.row() == cell->mergedRange.startRow &&
        index.column() == cell->mergedRange.startCol) {
        return QSize(cell->mergedRange.colSpan(), cell->mergedRange.rowSpan());
    }

    return QSize(1, 1);
}

QFont ReportDataModel::ensureFontAvailable(const QFont& requestedFont) const
{
    QFont font = requestedFont;

    // 检查字体是否在系统中可用
    QFontInfo fontInfo(font);
    QString actualFamily = fontInfo.family();
    QString requestedFamily = font.family();

    if (actualFamily != requestedFamily) {

        // 中文字体映射
        if (requestedFamily.contains("宋体") || requestedFamily == "SimSun") {
            QStringList alternatives = { "SimSun", "NSimSun", "宋体", "新宋体" };
            for (const QString& alt : alternatives) {
                font.setFamily(alt);
                QFontInfo altInfo(font);
                if (altInfo.family().contains(alt, Qt::CaseInsensitive)) {
                    return font;
                }
            }
        }

        if (requestedFamily.contains("黑体") || requestedFamily == "SimHei") {
            QStringList alternatives = { "Microsoft YaHei", "SimHei", "黑体" };
            for (const QString& alt : alternatives) {
                font.setFamily(alt);
                QFontInfo altInfo(font);
                if (altInfo.family().contains(alt, Qt::CaseInsensitive)) {
                    return font;
                }
            }
        }
        font.setFamily("");
    }

    return font;
}

QVariant ReportDataModel::data(const QModelIndex& index, int role) const
{
    if (!index.isValid())
        return QVariant();

    //  根据当前模式分发
    if (m_currentMode == HISTORY_MODE) {
        return getHistoryReportCellData(index, role);
    }
    else {
        return getRealtimeCellData(index, role);
    }
}

QVariant ReportDataModel::getRealtimeCellData(const QModelIndex& index, int role) const
{
    // 将原来data()函数中的实时模式逻辑全部移到这里
    const CellData* cell = getCell(index.row(), index.column());
    if (!cell)
        return QVariant();

    // 合并单元格处理
    if (cell->mergedRange.isValid() && cell->mergedRange.isMerged()) {
        bool isMainCell = (index.row() == cell->mergedRange.startRow &&
            index.column() == cell->mergedRange.startCol);

        switch (role) {
        case Qt::DisplayRole:
        case Qt::EditRole:
            return isMainCell ? cell->value.toString() : QVariant();
        case Qt::BackgroundRole:
            return QBrush(cell->style.backgroundColor);
        case Qt::ForegroundRole:
            return QBrush(cell->style.textColor);
        case Qt::FontRole: {
            QFont font = ensureFontAvailable(cell->style.font);
            return font;
        }
        case Qt::TextAlignmentRole:
            return static_cast<int>(cell->style.alignment);
        default:
            return QVariant();
        }
    }

    // 普通单元格处理
    switch (role) {
    case Qt::DisplayRole:
        return cell->displayText();
    case Qt::EditRole:
        return cell->editText();
    case Qt::BackgroundRole:
        return QBrush(cell->style.backgroundColor);
    case Qt::ForegroundRole:
        return QBrush(cell->style.textColor);
    case Qt::FontRole: {
        QFont font = ensureFontAvailable(cell->style.font);
        return font;
    }
    case Qt::TextAlignmentRole:
        return static_cast<int>(cell->style.alignment);
    default:
        return QVariant();
    }
}

QVariant ReportDataModel::getHistoryReportCellData(const QModelIndex& index, int role) const
{
    int row = index.row();
    int col = index.column();

    // 判断当前是显示配置还是报表数据
    if (m_fullTimeAxis.isEmpty()) {
        // ========== 显示配置文件内容（2列） ==========
		if (role == Qt::DisplayRole || role == Qt::EditRole) { // 只处理显示和编辑角色
            if (row >= m_historyConfig.columns.size()) return QVariant();

            const ReportColumnConfig& config = m_historyConfig.columns[row];
            if (col == 0) return config.displayName;
            if (col == 1) return config.rtuId;
        }

        if (role == Qt::BackgroundRole) { 
            return QBrush(QColor(250, 250, 250));
        }

        if (role == Qt::TextAlignmentRole) {
            return (int)(Qt::AlignVCenter | Qt::AlignLeft);
        }

    }
    else {
        // ========== 显示报表数据 ==========
        if (role == Qt::DisplayRole || role == Qt::EditRole) {
            // 表头行
            if (row == 0) {
                if (col == 0) return "时间";
                if (col - 1 < m_historyConfig.columns.size()) {
                    return m_historyConfig.columns[col - 1].displayName;
                }
                return QVariant();
            }

            // 数据行
            int dataRow = row - 1;
            if (dataRow >= 0 && dataRow < m_fullTimeAxis.size()) {
                if (col == 0) {
                    // 时间列
                    return m_fullTimeAxis[dataRow].toString("yyyy-MM-dd HH:mm:ss");
                }
                else if (col - 1 < m_historyConfig.columns.size()) {
                    // 数据列
                    QString rtuId = m_historyConfig.columns[col - 1].rtuId;
                    if (m_fullAlignedData.contains(rtuId)) {
                        const QVector<double>& values = m_fullAlignedData[rtuId];
                        if (dataRow < values.size()) {
                            double value = values[dataRow];
                            if (std::isnan(value) || std::isinf(value)) {
                                return "N/A";
                            }
                            return QString::number(value, 'f', 2);
                        }
                    }
                    return "N/A";
                }
            }
        }

        if (role == Qt::TextAlignmentRole) {
            return (int)(Qt::AlignVCenter | Qt::AlignLeft);
        }

        if (role == Qt::BackgroundRole) {
            if (row == 0) {
                return QBrush(QColor(220, 220, 220)); // 表头灰色
            }
            else {
                return (row % 2 == 0) ? QBrush(Qt::white) : QBrush(QColor(248, 248, 248)); // 斑马纹
            }
        }

        if (role == Qt::FontRole) {
            QFont font;
            if (row == 0) font.setBold(true);
            return font;
        }
    }

    return QVariant();
}

//  设置工作模式
void ReportDataModel::setWorkMode(WorkMode mode)
{
    if (m_currentMode != mode) {
        // 切换模式时清理对方模式的数据
        if (mode == HISTORY_MODE) {
            // 切换到报表模式，清理实时模式数据
            if (!m_cells.isEmpty()) {
                qDeleteAll(m_cells);
                m_cells.clear();
                qDebug() << "已清理实时模式数据";
            }
        }
        else {
            // 切换到实时模式，清理报表数据
            if (!m_fullTimeAxis.isEmpty() || !m_fullAlignedData.isEmpty()) {
                m_fullTimeAxis.clear();
                m_fullAlignedData.clear();
                m_historyConfig.columns.clear();
                m_historyConfig.reportName.clear();
                m_historyConfig.configFilePath.clear();
                m_reportName.clear();
                qDebug() << "已清理报表模式数据";
            }
        }

        m_currentMode = mode;
        qDebug() << "切换工作模式：" << (mode == REALTIME_MODE ? "实时模式" : "历史模式");
    }
}

//  检查是否有##绑定
bool ReportDataModel::hasDataBindings() const
{
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        if (it.value() && it.value()->isDataBinding) {
            return true;
        }
    }
    return false;
}

//  加载报表配置
bool ReportDataModel::loadReportConfig(const QString& filePath)
{
    QXlsx::Document xlsx(filePath);
    if (!xlsx.load()) {
        qWarning() << "无法打开配置文件:" << filePath;
        return false;
    }

    QXlsx::Worksheet* sheet = static_cast<QXlsx::Worksheet*>(xlsx.currentSheet());
    if (!sheet) return false;

    const QXlsx::CellRange range = sheet->dimension(); 
    if (!range.isValid()) return false;

    // 清空旧配置
    m_historyConfig.columns.clear();

    // 从文件名提取报表名称
    QFileInfo fileInfo(filePath);
    QString fileName = fileInfo.fileName();
    m_reportName = fileName.mid(6); // 去掉"#REPO_"
    m_reportName = m_reportName.left(m_reportName.lastIndexOf('.'));

    m_historyConfig.reportName = m_reportName;
    m_historyConfig.configFilePath = filePath;

    // 解析配置数据
    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        QVariant colA = sheet->read(row, 1); // A列
        QVariant colB = sheet->read(row, 2); // B列

        // 跳过空行
        if (colA.isNull() && colB.isNull()) continue;
        if (colA.toString().trimmed().isEmpty() &&
            colB.toString().trimmed().isEmpty()) continue;

        // 验证完整性
        if (colA.isNull() || colB.isNull()) {
            qWarning() << "配置文件第" << row << "行数据不完整";
            continue;
        }

        ReportColumnConfig colConfig; 
        colConfig.displayName = colA.toString().trimmed(); 
        colConfig.rtuId = colB.toString().trimmed();
        colConfig.sourceRow = row - 1;

        if (colConfig.rtuId.isEmpty()) {
            qWarning() << "配置文件第" << row << "行RTU号为空";
            continue;
        }

        m_historyConfig.columns.append(colConfig);
    }

    if (m_historyConfig.columns.isEmpty()) {
        qWarning() << "配置文件中没有有效的列配置";
        return false;
    }

    qDebug() << "成功加载报表配置：" << m_reportName
        << "，共" << m_historyConfig.columns.size() << "列";

    return true;
}

//  显示配置文件内容
void ReportDataModel::displayConfigFileContent()
{
    beginResetModel();

    clearAllCells();
    m_fullTimeAxis.clear();
    m_fullAlignedData.clear();

    int rowCount = m_historyConfig.columns.size();
    updateModelSize(rowCount, 2); // 2列：名称+RTU号

    endResetModel();
}

//  生成历史报表
void ReportDataModel::generateHistoryReport(
    const HistoryReportConfig& config,
    const QHash<QString, QVector<double>>& alignedData,
    const QVector<QDateTime>& timeAxis)
{
    beginResetModel();

    if (!m_cells.isEmpty()) {
        qDeleteAll(m_cells);
        m_cells.clear();
    }

    m_historyConfig = config;
    m_fullTimeAxis = timeAxis;
    m_fullAlignedData = alignedData;

    // 记录数据列的索引（时间列 + 所有RTU数据列）
    m_historyConfig.dataColumns.clear();
    m_historyConfig.dataColumns.insert(0);  // 时间列
    for (int i = 0; i < config.columns.size(); i++) {
        m_historyConfig.dataColumns.insert(i + 1);  // RTU数据列
    }

    int totalRows = timeAxis.size() + 1;
    int totalCols = config.columns.size() + 1;
    updateModelSize(totalRows, totalCols);

    endResetModel();

    qDebug() << "报表已生成，数据列索引：" << m_historyConfig.dataColumns;
}

//  生成时间轴（静态函数）
QVector<QDateTime> ReportDataModel::generateTimeAxis(const TimeRangeConfig& config)
{
    QVector<QDateTime> timeAxis;

    if (!config.isValid()) {
        qWarning() << "时间配置无效";
        return timeAxis;
    }

    //  新增：处理单点查询（间隔为0的情况）
    if (config.intervalSeconds == 0) {
        // 单点查询：只返回起始时间
        timeAxis.append(config.startTime);
        qDebug() << "生成单点查询时间：" << config.startTime.toString("yyyy-MM-dd HH:mm:ss");
        return timeAxis;
    }

    // 原有的时间范围查询逻辑
    QDateTime current = config.startTime;
    qint64 totalSeconds = config.startTime.secsTo(config.endTime);
    int estimatedSize = totalSeconds / config.intervalSeconds + 1;
    timeAxis.reserve(estimatedSize);

    while (current <= config.endTime) {
        timeAxis.append(current);
        current = current.addSecs(config.intervalSeconds);
    }

    qDebug() << "生成时间轴：" << timeAxis.size() << "个时间点";
    return timeAxis;
}

//  线性插值对齐（静态函数）
QHash<QString, QVector<double>> ReportDataModel::alignDataWithInterpolation(
    const QHash<QString, std::map<int64_t, std::vector<float>>>& rawData,
    const QVector<QDateTime>& timeAxis)
{
    QHash<QString, QVector<double>> result;

    for (auto it = rawData.constBegin(); it != rawData.constEnd(); ++it) {
        QString rtuId = it.key();
        const std::map<int64_t, std::vector<float>>& dataMap = it.value();

        QVector<double> alignedValues;
        alignedValues.reserve(timeAxis.size());

        if (dataMap.empty()) {
            alignedValues.fill(std::numeric_limits<double>::quiet_NaN(), timeAxis.size());
            result[rtuId] = alignedValues;
            continue;
        }

        for (const QDateTime& targetTime : timeAxis) {
            qint64 targetTs = targetTime.toSecsSinceEpoch();  //  直接用秒

            auto upper = dataMap.lower_bound(targetTs);

            //  判断是否完美匹配
            if (upper != dataMap.end() && upper->first == targetTs) {
                // 完美匹配，直接使用值，无需插值
                alignedValues.append(static_cast<double>(upper->second[0]));
                continue;
            }

            // 以下是原有的插值逻辑
            if (upper == dataMap.end()) {
                float lastValue = dataMap.rbegin()->second[0];
                alignedValues.append(static_cast<double>(lastValue));
            }
            else if (upper == dataMap.begin()) {
                float firstValue = upper->second[0];
                alignedValues.append(static_cast<double>(firstValue));
            }
            else {
                qint64 t2 = upper->first;
                float v2 = upper->second[0];

                auto lower = std::prev(upper);
                qint64 t1 = lower->first;
                float v1 = lower->second[0];

                if (t2 == t1) {
                    alignedValues.append(static_cast<double>(v1));
                }
                else {
                    // 线性插值
                    double ratio = static_cast<double>(targetTs - t1) / (t2 - t1);
                    double interpolated = v1 + (v2 - v1) * ratio;
                    alignedValues.append(interpolated);
                }
            }
        }

        result[rtuId] = alignedValues;
    }

    return result;
}

bool ReportDataModel::exportHistoryReportToExcel(
    const QString& fileName,
    QProgressDialog* progress)
{
    if (!hasHistoryData()) {
        qWarning() << "没有可导出的报表数据";
        return false;
    }

    QXlsx::Document xlsx;
    QXlsx::Worksheet* sheet = xlsx.currentWorksheet();

    int totalRows = m_fullTimeAxis.size();
    int totalCols = m_historyConfig.columns.size();

    // 1. 设置列宽
    sheet->setColumnWidth(1, 1, 20.0);  // 时间列
    for (int col = 2; col <= totalCols + 1; col++) {
        sheet->setColumnWidth(col, col, 15.0);
    }

    // 2. 写入表头
    QXlsx::Format headerFormat;
    headerFormat.setFontBold(true);
    headerFormat.setFontSize(11);
    headerFormat.setPatternBackgroundColor(QColor(220, 220, 220));
    headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    headerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    sheet->write(1, 1, "时间", headerFormat);
    for (int col = 0; col < totalCols; col++) {
        sheet->write(1, col + 2, m_historyConfig.columns[col].displayName, headerFormat);
    }

    if (progress) {
        progress->setValue(5);
        qApp->processEvents();
    }

    // 3. 写入数据
    QXlsx::Format dataFormat;
    dataFormat.setFontSize(10);
    dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    for (int row = 0; row < totalRows; row++) {
        if (progress && progress->wasCanceled()) {
            return false;
        }

        // 时间列
        QString timeStr = m_fullTimeAxis[row].toString("yyyy-MM-dd HH:mm:ss");
        sheet->write(row + 2, 1, timeStr, dataFormat);

        // 数据列
        for (int col = 0; col < totalCols; col++) {
            int modelRow = row + 1;  // +1 因为表头占第0行
            int modelCol = col + 1;  // +1 因为时间列占第0列

            //  优先检查用户是否编辑过这个单元格
            const CellData* cell = getCell(modelRow, modelCol);

            if (cell) {
                // 导出用户编辑的内容
                if (cell->hasFormula) {
                    QString fullFormula = cell->formula.startsWith('=') ?
                        cell->formula : ("=" + cell->formula);
                    sheet->write(row + 2, col + 2, fullFormula, dataFormat);
                }
                else {
                    sheet->write(row + 2, col + 2, cell->value, dataFormat);
                }
            }
            else {
                // 导出原始虚拟数据
                QString rtuId = m_historyConfig.columns[col].rtuId;

                if (m_fullAlignedData.contains(rtuId) &&
                    row < m_fullAlignedData[rtuId].size()) {
                    double value = m_fullAlignedData[rtuId][row];

                    if (std::isnan(value) || std::isinf(value)) {
                        sheet->write(row + 2, col + 2, "N/A", dataFormat);
                    }
                    else {
                        sheet->write(row + 2, col + 2, value, dataFormat);
                    }
                }
                else {
                    sheet->write(row + 2, col + 2, "N/A", dataFormat);
                }
            }
        }

        if (progress && row % 100 == 0) {
            int progressValue = 5 + (row * 90 / totalRows);
            progress->setValue(progressValue);
            qApp->processEvents();
        }
    }

    // 添加保存和返回
    if (progress) progress->setValue(100);

    if (!xlsx.saveAs(fileName)) {
        qWarning() << "保存Excel失败:" << fileName;
        return false;
    }

    return true;
}


bool ReportDataModel::setData(const QModelIndex& index, const QVariant& value, int role)
{
    if (!index.isValid() || role != Qt::EditRole)
        return false;

    if (m_currentMode == HISTORY_MODE) {
        if (m_fullTimeAxis.isEmpty()) {
            // ===== 配置阶段：直接修改配置数据 =====

            int row = index.row();
            int col = index.column();

            //  确保配置有足够的行（自动扩展，不触发信号）
            while (row >= m_historyConfig.columns.size()) {
                ReportColumnConfig newConfig;
                m_historyConfig.columns.append(newConfig);
            }

            // 同步更新模型行数（如果需要）
            if (row + 1 > m_maxRow) {
                m_maxRow = row + 1;
            }

            // 修改配置
            if (col == 0) {
                m_historyConfig.columns[row].displayName = value.toString().trimmed();
            }
            else if (col == 1) {
                m_historyConfig.columns[row].rtuId = value.toString().trimmed();
            }

            emit dataChanged(index, index, { role });
            return true;
        }
        else {
            // ===== 报表阶段：只允许编辑公式列 =====

            // 表头不可编辑
            if (index.row() == 0) {
                qWarning() << "表头不可编辑";
                return false;
            }

            // 数据列不可编辑
            if (m_historyConfig.dataColumns.contains(index.column())) {
                QMessageBox::warning(nullptr, "提示", "数据列为只读，请在右侧空白列添加公式。");
                return false;
            }

            // 公式列可编辑
            CellData* cell = ensureCell(index.row(), index.column());
            if (!cell) return false;

            QString text = value.toString();

            if (text.startsWith('=')) {
                cell->setFormula(text);
                calculateFormula(index.row(), index.column());
            }
            else {
                cell->value = value;
            }

            emit dataChanged(index, index, { role });
            return true;
        }
    }

    // ===== 实时模式编辑（保持原有逻辑）=====
    CellData* cell = ensureCell(index.row(), index.column());
    if (!cell) return false;

    QString text = value.toString();

    if (text.startsWith('##'))
    {
        cell->isDataBinding = true;
        cell->bindingKey = text;
        cell->hasFormula = false;
        cell->formula.clear();
        cell->value = "0";
        resolveDataBindings();
    }
    else if (text.startsWith('=')) {
        cell->isDataBinding = false;
        cell->setFormula(text);
        calculateFormula(index.row(), index.column());
    }
    else {
        cell->isDataBinding = false;
        if (cell->hasFormula) {
            cell->formula.clear();
            cell->hasFormula = false;
        }
        cell->value = value;
    }

    emit dataChanged(index, index, { role });
    emit cellChanged(index.row(), index.column());
    return true;
}

Qt::ItemFlags ReportDataModel::flags(const QModelIndex& index) const
{
    if (!index.isValid())
        return Qt::NoItemFlags;

    if (m_currentMode == HISTORY_MODE) {
        // ===== 判断当前是配置阶段还是报表阶段 =====
        if (m_fullTimeAxis.isEmpty()) {
            // 配置阶段：所有单元格可编辑
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
        }
        else {
            // 报表阶段：
            // - 表头只读
            // - 数据列只读
            // - 公式列可编辑
            if (index.row() == 0) {
                return Qt::ItemIsEnabled | Qt::ItemIsSelectable;  // 表头只读
            }

            //  判断是否为数据列
            if (m_historyConfig.dataColumns.contains(index.column())) {
                return Qt::ItemIsEnabled | Qt::ItemIsSelectable;  // 数据列只读
            }

            // 公式列可编辑
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
        }
    }

    // 检查是否是合并单元格的从属单元格
    const CellData* cell = getCell(index.row(), index.column());
    if (cell && cell->mergedRange.isValid() && cell->mergedRange.isMerged()) {
        // 如果不是主单元格，则不可编辑
        if (!(index.row() == cell->mergedRange.startRow &&
            index.column() == cell->mergedRange.startCol)) {
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
        }
    }

    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}
QVariant ReportDataModel::headerData(int section, Qt::Orientation orientation, int role) const
{
    if (role != Qt::DisplayRole)
        return QVariant();

    if (orientation == Qt::Horizontal) {
        QString colName;
        int temp = section;
        while (temp >= 0) {
            colName.prepend(QChar('A' + (temp % 26)));
            temp = temp / 26 - 1;
        }
        return colName;
    }
    else {
        return section + 1; // 行号从1开始
    }
}

// --- 行列操作实现 ---

bool ReportDataModel::insertRows(int row, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0) return false;

    beginInsertRows(QModelIndex(), row, row + count - 1);

    QHash<QPoint, CellData*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        CellData* cell = it.value();

        if (oldPos.x() >= row) {
            // 将此行及以下的单元格向下移动
            QPoint newPos(oldPos.x() + count, oldPos.y());
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startRow >= row) {
                    cell->mergedRange.startRow += count;
                    cell->mergedRange.endRow += count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越插入行的合并单元格信息
            if (cell->mergedRange.isValid() &&
                cell->mergedRange.startRow < row && cell->mergedRange.endRow >= row) {
                cell->mergedRange.endRow += count;
            }
        }
    }
    m_cells = newCells;
    m_maxRow += count;

    endInsertRows();
    return true;
}

bool ReportDataModel::removeRows(int row, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0 || row < 0 || row + count > m_maxRow)
            return false;

    beginRemoveRows(QModelIndex(), row, row + count - 1);

    QHash<QPoint, CellData*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        CellData* cell = it.value();

        if (oldPos.x() >= row && oldPos.x() < row + count) {
            // 删除被移除范围内的单元格
            delete cell;
        }
        else if (oldPos.x() >= row + count) {
            // 将被移除范围下方的单元格向上移动
            QPoint newPos(oldPos.x() - count, oldPos.y());
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startRow >= row + count) {
                    cell->mergedRange.startRow -= count;
                    cell->mergedRange.endRow -= count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越删除行的合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.endRow >= row + count) {
                    cell->mergedRange.endRow -= count;
                }
                else if (cell->mergedRange.endRow >= row) {
                    cell->mergedRange.endRow = row - 1;
                }

                // 如果合并范围无效了，清除合并信息
                if (cell->mergedRange.endRow < cell->mergedRange.startRow ||
                    cell->mergedRange.endCol < cell->mergedRange.startCol) {
                    cell->mergedRange = RTMergedRange();
                }
            }
        }
    }
    m_cells = newCells;
    m_maxRow -= count;

    endRemoveRows();
    return true;
}

bool ReportDataModel::insertColumns(int column, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0) return false;

    beginInsertColumns(QModelIndex(), column, column + count - 1);

    QHash<QPoint, CellData*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        CellData* cell = it.value();

        if (oldPos.y() >= column) {
            QPoint newPos(oldPos.x(), oldPos.y() + count);
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startCol >= column) {
                    cell->mergedRange.startCol += count;
                    cell->mergedRange.endCol += count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越插入列的合并单元格信息
            if (cell->mergedRange.isValid() &&
                cell->mergedRange.startCol < column && cell->mergedRange.endCol >= column) {
                cell->mergedRange.endCol += count;
            }
        }
    }
    m_cells = newCells;
    m_maxCol += count;

    endInsertColumns();
    return true;
}

bool ReportDataModel::removeColumns(int column, int count, const QModelIndex& parent)
{
    Q_UNUSED(parent)
        if (count <= 0 || column < 0 || column + count > m_maxCol)
            return false;

    beginRemoveColumns(QModelIndex(), column, column + count - 1);

    QHash<QPoint, CellData*> newCells;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        QPoint oldPos = it.key();
        CellData* cell = it.value();

        if (oldPos.y() >= column && oldPos.y() < column + count) {
            delete cell;
        }
        else if (oldPos.y() >= column + count) {
            QPoint newPos(oldPos.x(), oldPos.y() - count);
            newCells[newPos] = cell;

            // 更新合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.startCol >= column + count) {
                    cell->mergedRange.startCol -= count;
                    cell->mergedRange.endCol -= count;
                }
            }
        }
        else {
            newCells[oldPos] = cell;

            // 更新跨越删除列的合并单元格信息
            if (cell->mergedRange.isValid()) {
                if (cell->mergedRange.endCol >= column + count) {
                    cell->mergedRange.endCol -= count;
                }
                else if (cell->mergedRange.endCol >= column) {
                    cell->mergedRange.endCol = column - 1;
                }

                // 如果合并范围无效了，清除合并信息
                if (cell->mergedRange.endRow < cell->mergedRange.startRow ||
                    cell->mergedRange.endCol < cell->mergedRange.startCol) {
                    cell->mergedRange = RTMergedRange();
                }
            }
        }
    }
    m_cells = newCells;
    m_maxCol -= count;

    endRemoveColumns();
    return true;
}

// --- 文件操作实现 ---

bool ReportDataModel::loadFromExcel(const QString& fileName)
{
    beginResetModel();
    clearAllCells(); // 这会清空 m_cells 和尺寸向量
    m_maxRow = 100;    // 恢复默认行数
    m_maxCol = 26;     // 恢复默认列数
    endResetModel();

    // 重新发出 beginResetModel 信号，为加载新数据做准备
    beginResetModel();
    bool result = ExcelHandler::loadFromFile(fileName, this);
    endResetModel();

    // 加载成功后，统一重新计算所有公式
    if (result) {
        resolveDataBindings();
        recalculateAllFormulas();
    }
    return result;
}

// 这个函数是核心，它负责收集所有需要绑定的Key，
// 并通过信号发送给外部（例如MainWindow）去查询数据。
void ReportDataModel::resolveDataBindings()
{
    QList<QString> keysToResolve;
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->isDataBinding) {
            keysToResolve.append(cell->bindingKey);
        }
    }

    if (keysToResolve.isEmpty()) {
        return;
    }

    QHash<QString, QVariant> resolvedData = UniversalQueryEngine::instance().queryValuesForBindingKeys(keysToResolve);

    // 更新单元格的值
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->isDataBinding && resolvedData.contains(cell->bindingKey)) {
            cell->value = resolvedData[cell->bindingKey];
        }
    }

    // 通知整个视图刷新，因为多个单元格数据可能已改变
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
}

bool ReportDataModel::saveToExcel(const QString& fileName)
{
    // 直接委托给ExcelHandler处理
    return ExcelHandler::saveToFile(fileName, this);
}


void ReportDataModel::clearAllCells()
{
    if (m_cells.isEmpty()) return;

    // 注意：这里不调用begin/endResetModel，因为调用方(loadFromExcel)会负责
    qDeleteAll(m_cells);
    m_cells.clear();
    clearSizes(); // <-- 新增这一行
}

void ReportDataModel::addCellDirect(int row, int col, CellData* cell)
{
    // 此方法专为Excel高速加载设计，不触发信号
    QPoint key(row, col);
    if (m_cells.contains(key)) {
        delete m_cells.take(key);
    }
    m_cells.insert(key, cell);
}

void ReportDataModel::updateModelSize(int newRowCount, int newColCount)
{
    // 更新内部行列数，保证最小尺寸
    m_maxRow = qMax(100, newRowCount);
    m_maxCol = qMax(26, newColCount);
    // 注意：同样不调用begin/endResetModel，由调用方负责
}

const QHash<QPoint, CellData*>& ReportDataModel::getAllCells() const
{
    return m_cells;
}

void ReportDataModel::recalculateAllFormulas()
{
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        if (it.value() && it.value()->hasFormula) {
            const QPoint& pos = it.key();
            calculateFormula(pos.x(), pos.y());
        }
    }
    // 计算完成后，通知视图全局刷新数据
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
}

// --- 工具和私有方法实现 ---

QString ReportDataModel::cellAddress(int row, int col) const
{
    QString colName;
    int temp = col;
    while (temp >= 0) {
        colName.prepend(QChar('A' + (temp % 26)));
        temp = temp / 26 - 1;
    }
    return colName + QString::number(row + 1);
}

void ReportDataModel::calculateFormula(int row, int col)
{
    CellData* cell = getCell(row, col);
    if (!cell || !cell->hasFormula)
        return;

    // 调用公式引擎计算结果
    QVariant result = m_formulaEngine->evaluate(cell->formula, this, row, col);
    cell->value = result;
}

CellData* ReportDataModel::getCell(int row, int col)
{
    return m_cells.value(QPoint(row, col), nullptr);
}

const CellData* ReportDataModel::getCell(int row, int col) const
{
    return m_cells.value(QPoint(row, col), nullptr);
}

CellData* ReportDataModel::ensureCell(int row, int col)
{
    QPoint key(row, col);
    if (!m_cells.contains(key)) {
        // 如果单元格不存在，则创建一个新的
        m_cells[key] = new CellData();
    }
    return m_cells[key];
}

void ReportDataModel::setRowHeight(int row, double height)
{
    if (row >= m_rowHeights.size()) {
        m_rowHeights.resize(row + 1);
    }
    m_rowHeights[row] = height;
}

double ReportDataModel::getRowHeight(int row) const
{
    if (row < m_rowHeights.size()) {
        return m_rowHeights[row];
    }
    return 0; // 返回0表示使用默认值
}

void ReportDataModel::setColumnWidth(int col, double width)
{
    if (col >= m_columnWidths.size()) {
        m_columnWidths.resize(col + 1);
    }
    m_columnWidths[col] = width;
}

double ReportDataModel::getColumnWidth(int col) const
{
    if (col < m_columnWidths.size()) {
        return m_columnWidths[col];
    }
    return 0; // 返回0表示使用默认值
}

const QVector<double>& ReportDataModel::getAllRowHeights() const
{
    return m_rowHeights;
}

const QVector<double>& ReportDataModel::getAllColumnWidths() const
{
    return m_columnWidths;
}

void ReportDataModel::clearSizes()
{
    m_rowHeights.clear();
    m_columnWidths.clear();
}

void ReportDataModel::setGlobalConfig(const GlobalDataConfig& config)
{
    m_globalConfig = config;
}

GlobalDataConfig ReportDataModel::getGlobalConfig() const
{
    return m_globalConfig;
}

void ReportDataModel::updateGlobalTimeRange(const TimeRangeConfig& timeRange)
{
    m_globalConfig.globalTimeRange = timeRange;
    // 可选：发射信号通知时间范围已更新
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
}

void ReportDataModel::restoreBindingsToConfigStage()
{
    if (m_currentMode != REALTIME_MODE) {
        qWarning() << "还原配置仅支持实时模式";
        return;
    }

    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->isDataBinding) {
            // 将数值恢复为 ##RTU号
            cell->value = cell->bindingKey;
        }
    }

    // 通知视图刷新所有单元格
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));

    qDebug() << "实时模式配置已还原";
}
