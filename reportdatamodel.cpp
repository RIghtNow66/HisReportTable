#include "reportdatamodel.h"
#include "formulaengine.h"
#include "excelhandler.h" // 用于文件操作
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

#include "UniversalQueryEngine.h"
#include "DayReportParser.h"
// QXlsx相关（检查是否已包含）
#include "xlsxdocument.h"      // 用于 QXlsx::Document
#include "xlsxcellrange.h"     // 用于 QXlsx::CellRange
#include "xlsxworksheet.h"     // 用于 QXlsx::Worksheet
#include "xlsxformat.h"        // 用于 QXlsx::Format


ReportDataModel::ReportDataModel(QObject* parent)
    : QAbstractTableModel(parent)
    , m_maxRow(100) // 默认初始行数
    , m_maxCol(26)  // 默认初始列数 (A-Z)
    , m_formulaEngine(new FormulaEngine(this))
    , m_reportType(NORMAL_EXCEL) // 默认是普通Excel
    , m_dayParser(nullptr)
{
}

ReportDataModel::~ReportDataModel()
{
    clearAllCells();
}

// ===== 统一的模板加载入口 =====
bool ReportDataModel::loadReportTemplate(const QString& fileName)
{
    // 1. 清理旧状态，并将刷新标记重置为 false
    clearAllCells();

    // 2. 加载基础Excel数据
    if (!loadFromExcelFile(fileName)) {
        return false;
    }

    // 3. 根据文件名判断报表类型
    QFileInfo fileInfo(fileName);
    QString baseName = fileInfo.fileName();
    m_reportName = fileInfo.baseName();

    if (baseName.startsWith("##Day_", Qt::CaseInsensitive)) {
        m_reportType = DAY_REPORT;
        qDebug() << "检测到日报模板，开始解析...";
        m_dayParser = new DayReportParser(this, this);
        // scanAndParse 会识别模板中的标记
        if (!m_dayParser->scanAndParse()) {
            // 如果解析失败（例如缺少#Date），清理并报错
            clearAllCells();
            return false;
        }
    }
    else if (baseName.startsWith("##Month_", Qt::CaseInsensitive)) {
        m_reportType = MONTH_REPORT;
        qDebug() << "检测到月报模板 (功能待实现)...";
        // 预留月报解析器逻辑
    }
    else {
        m_reportType = NORMAL_EXCEL;
        qDebug() << "加载普通Excel文件。";
        // 对于普通Excel，在加载阶段我们只识别 ## 绑定，不计算公式。
        // 公式计算将在用户点击“刷新数据”时与数据绑定一起刷新。
        resolveDataBindings();
    }

    // 无论加载何种类型，都通知视图进行一次初始显示
    notifyDataChanged();
    return true;
}


void ReportDataModel::recalculateAllFormulas()
{
    qDebug() << "开始增量计算公式...";
    int calculatedCount = 0;
    int skippedCount = 0;

    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->hasFormula) {
            const QPoint& pos = it.key();

            // ===== 【核心逻辑】只计算未计算过的公式 =====
            if (cell->formulaCalculated) {
                skippedCount++;
                continue;  // 跳过已计算的公式
            }
            // ===== 【核心逻辑结束】 =====

            calculateFormula(pos.x(), pos.y());
            calculatedCount++;
        }
    }

    qDebug() << QString("公式计算完成: 计算 %1 个, 跳过 %2 个")
        .arg(calculatedCount).arg(skippedCount);

    notifyDataChanged();
}

// ===== 新增：统一的数据刷新入口 =====
bool ReportDataModel::refreshReportData(QProgressDialog* progress)
{
    bool completedSuccessfully = false;

    switch (m_reportType) {
    case DAY_REPORT:
        if (m_dayParser) {
            // 先执行数据查询
            completedSuccessfully = m_dayParser->executeQueries(progress);

            // ===== 如果查询没有被取消，则立即重新计算所有公式 =====
            if (completedSuccessfully) {
                qDebug() << "数据查询完成，开始计算延迟公式...";
                recalculateAllFormulas();
            }
        }
        break;

    case MONTH_REPORT:
        qDebug() << "月报刷新功能待实现";
        break;

    case NORMAL_EXCEL:
        resolveDataBindings();
        // ===== 刷新数据绑定后，计算所有公式 =====
        recalculateAllFormulas();
        completedSuccessfully = true;
        break;
    }
    return completedSuccessfully;
}

void ReportDataModel::restoreToTemplate()
{
    if (m_reportType == DAY_REPORT && m_dayParser) {
        m_dayParser->restoreToTemplate();
        notifyDataChanged(); // 通知视图刷新
    }
	// 对于普通Excel，恢复公式单元格显示公式文本
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->hasFormula) {
            cell->formulaCalculated = false;  // 重置为未计算
            cell->value = cell->formula;       // 恢复显示公式文本
        }
    }
}

ReportDataModel::ReportType ReportDataModel::getReportType() const {
    return m_reportType;
}

DayReportParser* ReportDataModel::getDayParser() const {
    return m_dayParser;
}

void ReportDataModel::notifyDataChanged() {
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
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
    if (!index.isValid()) return QVariant();
    const CellData* cell = getCell(index.row(), index.column());
    if (!cell) return QVariant();

    switch (role) {
    case Qt::DisplayRole:
        return cell->displayText();
    case Qt::EditRole:
        return cell->editText();
    case Qt::BackgroundRole:
        if (cell->cellType == CellData::DataMarker && cell->queryExecuted && !cell->querySuccess) {
            return QBrush(QColor(255, 220, 220)); // 查询失败显示淡红色
        }
        return QBrush(cell->style.backgroundColor);
    case Qt::ForegroundRole:
        return QBrush(cell->style.textColor);
    case Qt::FontRole:
        return ensureFontAvailable(cell->style.font);
    case Qt::TextAlignmentRole:
        return static_cast<int>(cell->style.alignment);
    default:
        return QVariant();
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

bool ReportDataModel::setData(const QModelIndex& index, const QVariant& value, int role)
{
    if (!index.isValid() || role != Qt::EditRole) {
        return false;
    }

    CellData* cell = ensureCell(index.row(), index.column());
    if (!cell) {
        return false;
    }

    QString text = value.toString();

    if (cell->hasFormula && text == cell->formula) {
        return false;  // 公式未修改，不做任何处理
    }

    if (cell->isMarker() && text == cell->originalMarker) {
        return false;
    }

    // 清理旧的状态
    cell->isDataBinding = false;
    cell->hasFormula = false;
    cell->formulaCalculated = false;  // 重置计算标记
    cell->cellType = CellData::NormalCell;
    cell->originalMarker.clear();
    cell->formula.clear();
    cell->bindingKey.clear();

    // ===== 只支持 #=# 延迟公式 =====
    if (text.startsWith("#=#")) {
        cell->hasFormula = true;
        cell->formula = text;
        cell->value = text;  // 显示公式文本
        cell->formulaCalculated = false;  // 标记为未计算
        // 不立即调用 calculateFormula
    }
    else if (text.startsWith("##")) {
        cell->isDataBinding = true;
        cell->bindingKey = text;
        cell->value = "0";
        if (m_reportType == NORMAL_EXCEL) {
            resolveDataBindings();
        }
    }
    else {
        // 普通文本或其他标记
        cell->value = value;

        if (text.startsWith("#d#", Qt::CaseInsensitive)) {
            cell->cellType = CellData::DataMarker;
            cell->originalMarker = text;
            cell->queryExecuted = false;
            cell->querySuccess = false;
        }
        else if (text.startsWith("#t#", Qt::CaseInsensitive)) {
            cell->cellType = CellData::TimeMarker;
            cell->originalMarker = text;
        }
        else if (text.startsWith("#Date", Qt::CaseInsensitive)) {
            cell->cellType = CellData::DateMarker;
            cell->originalMarker = text;
        }
    }

    emit dataChanged(index, index, { role });
    emit cellChanged(index.row(), index.column());
    return true;
}


Qt::ItemFlags ReportDataModel::flags(const QModelIndex& index) const
{
    if (!index.isValid()) return Qt::NoItemFlags;

    const CellData* cell = getCell(index.row(), index.column());
    if (cell && cell->mergedRange.isMerged()) {
        // 如果不是合并单元格的主单元格，则只启用，不可选择/编辑
        if (index.row() != cell->mergedRange.startRow || index.column() != cell->mergedRange.startCol) {
            return Qt::ItemIsEnabled;
        }
    }

    // 所有主单元格或非合并单元格都可编辑
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

bool ReportDataModel::loadFromExcelFile(const QString& fileName)
{
    beginResetModel();
    // 清理工作已移至 loadReportTemplate
    m_maxRow = 100;
    m_maxCol = 26;
    endResetModel();

    beginResetModel();
    bool result = ExcelHandler::loadFromFile(fileName, this);
    endResetModel();

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
    if (m_cells.isEmpty() && !m_dayParser) return;

    beginResetModel();
    qDeleteAll(m_cells);
    m_cells.clear();
    clearSizes();

    // 清理所有解析器
    delete m_dayParser;
    m_dayParser = nullptr;
    // delete m_monthParser; m_monthParser = nullptr;

    m_reportType = NORMAL_EXCEL;
    m_reportName.clear();
    endResetModel();
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
    cell->formulaCalculated = true;  // 标记已计算
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
