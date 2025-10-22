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
#include <QRegularExpression>

#include "reportdatamodel.h"
#include "formulaengine.h"
#include "excelhandler.h" // 用于文件操作
#include "BaseReportParser.h"
#include "MonthReportParser.h"
#include "DayReportParser.h"
#include "UnifiedQueryParser.h"

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
    , m_parser(nullptr)
    , m_editMode(true)
    , m_currentMode(TEMPLATE_MODE)      // 默认模板模式
    , m_templateType(TemplateType::DAY_REPORT)        // 默认日报
{
}

ReportDataModel::~ReportDataModel()
{
    clearAllCells();
    delete m_parser;  // 新增：释放解析器
    m_parser = nullptr;
}

// ===== 统一的模板加载入口 =====
bool ReportDataModel::loadReportTemplate(const QString& fileName)
{
    // 1. 清理旧状态
    clearAllCells();

    // 2. 加载基础Excel数据
    if (!loadFromExcelFile(fileName)) {
        return false;
    }

    // 3. 根据文件名判断报表类型
    QFileInfo fileInfo(fileName);
    QString baseName = fileInfo.fileName();
    m_reportName = fileInfo.baseName();

    qDebug() << "加载文件：" << baseName;

    // ===== 识别统一查询模式 =====
    if (baseName.startsWith("##REPO_", Qt::CaseInsensitive)) {
        qDebug() << "检测到统一查询模式文件：" << baseName;
        m_currentMode = UNIFIED_QUERY_MODE;  // 明确设置
        m_reportType = NORMAL_EXCEL;

        bool success = loadUnifiedQueryConfig(fileName);
        if (success) {
            setEditMode(true);
            notifyDataChanged();
            qDebug() << "统一查询配置加载成功，当前模式：" << m_currentMode;
        }
        return success;
    }

    // ===== 识别日报模式 =====
    if (baseName.startsWith("##Day_", Qt::CaseInsensitive)) {
        m_reportType = DAY_REPORT;
        m_currentMode = TEMPLATE_MODE;  // 明确设置
        qDebug() << "检测到日报模板，开始解析...";

        m_parser = new DayReportParser(this, this);

        if (!m_parser->scanAndParse()) {
            clearAllCells();
            return false;
        }

        // 连接预查询完成信号
        connect(m_parser, &BaseReportParser::prefetchCompleted,
            this, [this](bool hasData, int dataCount, int successCount, int totalCount) {
                // 预查询完成后的处理在 MainWindow 中
            }, Qt::QueuedConnection);

        setEditMode(true);
        notifyDataChanged();
        return true;
    }

    // ===== 识别月报模式 =====
    if (baseName.startsWith("##Month_", Qt::CaseInsensitive)) {
        m_reportType = MONTH_REPORT;
        m_currentMode = TEMPLATE_MODE;  //  明确设置
        qDebug() << "检测到月报模板，开始解析...";

        m_parser = new MonthReportParser(this, this);

        if (!m_parser->scanAndParse()) {
            clearAllCells();
            return false;
        }

        // 连接预查询完成信号
        connect(m_parser, &BaseReportParser::prefetchCompleted,
            this, [this](bool hasData, int dataCount, int successCount, int totalCount) {
                // 预查询完成后的处理在 MainWindow 中
            }, Qt::QueuedConnection);

        setEditMode(true);
        notifyDataChanged();
        return true;
    }

    // ===== 普通 Excel =====
    m_reportType = NORMAL_EXCEL;
    m_currentMode = TEMPLATE_MODE;  //  明确设置
    qDebug() << "加载普通Excel文件。";

    setEditMode(true);
    notifyDataChanged();
    return true;
}

// ===== 检查公式的依赖单元格是否已计算 =====
bool ReportDataModel::checkFormulaDependenciesReady(int row, int col)
{
    QSet<QPoint> visited;
    if (detectCircularDependency(row, col, visited)) {
        qWarning() << QString("检测到循环依赖: 单元格(%1, %2)").arg(row).arg(col);
        return false;
    }

    const CellData* cell = getCell(row, col);
    if (!cell || !cell->hasFormula) {
        return true;
    }

    QString formula = cell->formula;
    QRegularExpression cellRefRegex(R"([A-Z]+\d+)");
    QRegularExpressionMatchIterator it = cellRefRegex.globalMatch(formula);

    while (it.hasNext()) {
        QRegularExpressionMatch match = it.next();
        QString cellRef = match.captured(0);
        QPoint refPos = parseCellReference(cellRef);

        if (refPos.x() == -1 || refPos.y() == -1) {
            continue;
        }

        const CellData* refCell = getCell(refPos.x(), refPos.y());
        if (refCell && refCell->hasFormula && !refCell->formulaCalculated) {
            return false;
        }
    }

    return true;
}

bool ReportDataModel::detectCircularDependency(int row, int col, QSet<QPoint>& visitedCells)
{
    QPoint current(row, col);

    if (visitedCells.contains(current)) {
        return true;  // 检测到循环
    }

    const CellData* cell = getCell(row, col);
    if (!cell || !cell->hasFormula) {
        return false;
    }

    visitedCells.insert(current);

    QString formula = cell->formula;
    QRegularExpression cellRefRegex(R"([A-Z]+\d+)");
    QRegularExpressionMatchIterator it = cellRefRegex.globalMatch(formula);

    while (it.hasNext()) {
        QRegularExpressionMatch match = it.next();
        QString cellRef = match.captured(0);
        QPoint refPos = parseCellReference(cellRef);

        if (refPos.x() == -1 || refPos.y() == -1) {
            continue;
        }

        if (detectCircularDependency(refPos.x(), refPos.y(), visitedCells)) {
            return true;
        }
    }

    visitedCells.remove(current);
    return false;
}

// ===== 解析单元格引用（如 "I5" -> QPoint(4, 8)） =====
QPoint ReportDataModel::parseCellReference(const QString& cellRef) const
{
    QRegularExpression regex(R"(([A-Z]+)(\d+))");
    QRegularExpressionMatch match = regex.match(cellRef);

    if (!match.hasMatch()) {
        return QPoint(-1, -1);
    }

    QString colStr = match.captured(1);
    int rowNum = match.captured(2).toInt();

    // 转换列名为列索引（A=0, B=1, ..., Z=25, AA=26, ...）
    int col = 0;
    for (int i = 0; i < colStr.length(); ++i) {
        col = col * 26 + (colStr[i].unicode() - 'A' + 1);
    }
    col -= 1;  // 转换为0基索引

    return QPoint(rowNum - 1, col);  // 转换为0基索引
}


void ReportDataModel::recalculateAllFormulas()
{
    qDebug() << "开始增量计算公式...";

    // 如果脏标记为空，说明是首次计算或全量刷新
    if (m_dirtyFormulas.isEmpty()) {
        qDebug() << "脏标记为空，标记所有公式";
        for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
            CellData* cell = it.value();
            if (cell && cell->hasFormula) {
                cell->formulaCalculated = false;
                m_dirtyFormulas.insert(it.key());
            }
        }
    }

    if (m_dirtyFormulas.isEmpty()) {
        qDebug() << "没有需要计算的公式";
        return;
    }

    qDebug() << QString("待计算公式总数: %1").arg(m_dirtyFormulas.size());

    int maxIterations = 5;
    int iteration = 0;

    while (iteration < maxIterations && !m_dirtyFormulas.isEmpty()) {
        int calculatedCount = 0;
        QSet<QPoint> stillDirty;

        for (const QPoint& pos : m_dirtyFormulas) {
            CellData* cell = getCell(pos.x(), pos.y());

            if (!cell || !cell->hasFormula) {
                continue;
            }

            if (cell->formulaCalculated) {
                continue;
            }

            // 检查依赖是否已就绪
            if (checkFormulaDependenciesReady(pos.x(), pos.y())) {
                calculateFormula(pos.x(), pos.y());
                calculatedCount++;
            }
            else {
                stillDirty.insert(pos);
            }
        }

        m_dirtyFormulas = stillDirty;

        qDebug() << QString("第 %1 轮计算: 计算 %2 个, 剩余 %3 个")
            .arg(iteration + 1).arg(calculatedCount).arg(m_dirtyFormulas.size());

        if (calculatedCount == 0) {
            if (!m_dirtyFormulas.isEmpty()) {
                qWarning() << QString("检测到 %1 个公式无法计算，可能存在循环依赖").arg(m_dirtyFormulas.size());
            }
            break;
        }

        iteration++;
    }

    if (iteration >= maxIterations && !m_dirtyFormulas.isEmpty()) {
        qWarning() << "公式计算达到最大迭代次数";
    }

    m_dirtyFormulas.clear();
    notifyDataChanged();
}

//bool ReportDataModel::refreshReportData(QProgressDialog* progress)
//{
//    // ===== 【步骤1】检测变化类型 =====
//    ChangeType changeType = detectChanges();
//
//    QString changeMsg;
//    bool needQuery = false;
//
//    switch (changeType) {
//    case NO_CHANGE:
//        changeMsg = "数据已是最新，无需刷新。";
//        QMessageBox::information(nullptr, "无需刷新", changeMsg);
//        saveRefreshSnapshot();
//        return true;
//
//    case FORMULA_ONLY: {
//        int newFormulaCount = 0;
//        for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
//            if (it.value() && it.value()->hasFormula && !it.value()->formulaCalculated) {
//                newFormulaCount++;
//            }
//        }
//        changeMsg = QString("检测到 %1 个新增公式，是否计算？\n\n注意：不会重新查询数据。")
//            .arg(newFormulaCount);
//        needQuery = false;
//        break;
//    }
//
//    case BINDING_ONLY: {
//        int newMarkerCount = 0;
//        if (m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) {
//            for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
//                const CellData* cell = it.value();
//                if (cell && cell->cellType == CellData::DataMarker && !cell->queryExecuted) {
//                    newMarkerCount++;
//                }
//            }
//        }
//        else {
//            QList<QString> newBindings = getNewBindings();
//            newMarkerCount = newBindings.size();
//        }
//
//        changeMsg = QString("检测到 %1 个新增数据标记，是否查询？\n\n将查询新数据并重新计算所有公式。")
//            .arg(newMarkerCount);
//        needQuery = true;
//        break;
//    }
//
//    case MIXED_CHANGE: {
//        if (m_isFirstRefresh) {
//            needQuery = true;
//        }
//        else {
//            int newFormulaCount = 0;
//            for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
//                if (it.value() && it.value()->hasFormula && !it.value()->formulaCalculated) {
//                    newFormulaCount++;
//                }
//            }
//
//            int newDataCount = 0;
//            if (m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) {
//                for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
//                    const CellData* cell = it.value();
//                    if (cell && cell->cellType == CellData::DataMarker && !cell->queryExecuted) {
//                        newDataCount++;
//                    }
//                }
//            }
//            else {
//                newDataCount = getNewBindings().size();
//            }
//
//            changeMsg = QString("检测到数据变更：\n"
//                "• 新增数据标记: %1 个\n"
//                "• 新增公式: %2 个\n\n"
//                "是否执行完整刷新？")
//                .arg(newDataCount)
//                .arg(newFormulaCount);
//            needQuery = true;
//        }
//        break;
//    }
//    }
//
//    // ===== 【步骤2】用户确认（首次刷新跳过） =====
//    if (!m_isFirstRefresh && !changeMsg.isEmpty()) {
//        auto reply = QMessageBox::question(nullptr, "确认刷新", changeMsg,
//            QMessageBox::Yes | QMessageBox::No,
//            QMessageBox::Yes);
//        if (reply == QMessageBox::No) {
//            return false;
//        }
//    }
//
//    if (needQuery && (m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) && m_parser) {
//        if (!m_parser->isCacheValid()) {
//            qDebug() << "缓存已过期，将重新查询";
//            m_parser->invalidateCache();
//        }
//    }
//
//    // ===== 【步骤3】执行刷新操作 =====
//    bool completedSuccessfully = false;
//    int actualSuccessCount = 0;
//
//    // 根据 needQuery 决定是否查询
//    if (needQuery) {
//        // 需要查询数据
//        if ((m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) && m_parser) {
//            completedSuccessfully = m_parser->executeQueries(progress);
//
//            // 统计实际成功的数据点
//            for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
//                const CellData* cell = it.value();
//                if (cell && cell->cellType == CellData::DataMarker &&
//                    cell->queryExecuted && cell->querySuccess) {
//                    actualSuccessCount++;
//                }
//            }
//        }
//        else if (m_reportType == NORMAL_EXCEL) {
//            qDebug() << "普通Excel不支持数据查询";
//            completedSuccessfully = true;
//        }
//    }
//    else {
//        // 不需要查询，直接标记为成功
//        completedSuccessfully = true;
//    }
//
//    // 在查询完成后添加
//    if (completedSuccessfully) {
//        qDebug() << "数据填充完成，开始计算公式...";
//        recalculateAllFormulas();
//
//        // 执行内存优化
//        optimizeMemory();
//    }
//
//    // ===== 【修改】在数据填充完成后，再计算公式 =====
//    qDebug() << "数据填充完成，开始计算公式...";
//    recalculateAllFormulas();
//    
//    optimizeMemory();
//
//    bool shouldEnterRunMode = false;
//
//    // 保存快照
//    if ((m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) && m_parser) {
//        int totalTasks = m_parser->getPendingQueryCount();
//        if (totalTasks > 0) {
//            double successRate = static_cast<double>(actualSuccessCount) / totalTasks;
//            if (successRate >= 0.5) {
//                saveRefreshSnapshot();
//                shouldEnterRunMode = true;  // 进入运行模式
//                qDebug() << QString("保存快照: 成功率 %1%").arg(successRate * 100, 0, 'f', 1);
//            }
//            else {
//                shouldEnterRunMode = false;  // 保持编辑模式，允许用户修改重试
//                qWarning() << QString("成功率过低 (%1%)，不保存快照").arg(successRate * 100, 0, 'f', 1);
//            }
//        }
//        else {
//            saveRefreshSnapshot();
//            shouldEnterRunMode = true;
//        }
//    }
//    else {
//        saveRefreshSnapshot();
//        shouldEnterRunMode = true;
//    }
//
//    if (shouldEnterRunMode) {
//        setEditMode(false);
//    }
//
//    return completedSuccessfully;
//}

bool ReportDataModel::refreshReportData(QProgressDialog* progress)
{
    // ===== 新增：根据模式分发 =====
    if (m_currentMode == TEMPLATE_MODE) {
        return refreshTemplateReport(progress);
    }
    else if (m_currentMode == UNIFIED_QUERY_MODE) {
        return refreshUnifiedQuery(progress);
    }

    return false;
}

void ReportDataModel::restoreTemplateReport()
{
    qDebug() << "还原模板模式配置...";

    if (m_parser) {
        m_parser->restoreToTemplate();
    }

    // 重置所有公式为未计算状态
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->hasFormula) {
            cell->formulaCalculated = false;
            cell->value = cell->formula;
        }
    }

    // 清空快照
    m_lastSnapshot.bindingKeys.clear();
    m_lastSnapshot.formulaCells.clear();
    m_lastSnapshot.dataMarkerCells.clear();
    m_isFirstRefresh = true;

    m_dirtyFormulas.clear();

    // 恢复编辑状态
    if (m_parser) {
        m_parser->setEditState(BaseReportParser::CONFIG_EDIT);
    }
    setEditMode(true);

    qDebug() << "模板模式已还原到配置状态";
    notifyDataChanged();
}

void ReportDataModel::restoreUnifiedQuery()
{
    qDebug() << "还原统一查询配置...";

    // ===== 检查是否有用户添加的公式 =====
    bool hasUserFormulas = false;
    int formulaCount = 0;

    for (int row = 0; row < m_maxRow; ++row) {
        for (int col = m_dataColumnCount + 1; col < m_maxCol; ++col) {
            const CellData* cell = getCell(row, col);
            if (cell && cell->hasFormula) {
                hasUserFormulas = true;
                formulaCount++;
            }
        }
    }

    if (hasUserFormulas) {
        qDebug() << QString("检测到 %1 个用户公式").arg(formulaCount);

        // ===== 弹出警告对话框 =====
        auto reply = QMessageBox::warning(nullptr, "确认还原配置",
            QString("检测到您添加了 %1 个自定义公式/列。\n\n"
                "还原配置将清除所有查询数据和自定义公式。\n\n"
                "是否继续？").arg(formulaCount),
            QMessageBox::Yes | QMessageBox::No,
            QMessageBox::No);

        if (reply == QMessageBox::No) {
            qDebug() << "用户取消还原";
            return;
        }
    }

    // ===== 执行还原 =====
    if (m_parser) {
        m_parser->restoreToTemplate();
    }

    // 清除所有用户添加的单元格（包括公式列）
    QList<QPoint> toRemove;
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        int col = it.key().y();
        // 删除数据列之外的所有单元格
        if (col > 1) {  // 配置阶段只有前2列
            toRemove.append(it.key());
        }
    }

    for (const QPoint& pos : toRemove) {
        delete m_cells.take(pos);
    }

    // 恢复为配置视图（2列）
    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    if (queryParser) {
        const HistoryReportConfig& config = queryParser->getConfig();
        beginResetModel();
        updateModelSize(config.columns.size(), 2);
        m_dataColumnCount = 0;  // 清空数据列计数
        endResetModel();

        m_dirtyFormulas.clear();

        qDebug() << "统一查询已还原到配置阶段";
        notifyDataChanged();

        QMessageBox::information(nullptr, "还原完成", "已还原到配置文件状态。");
    }
}
void ReportDataModel::restoreToTemplate()
{
    if (m_currentMode == TEMPLATE_MODE) {
        restoreTemplateReport();
    }
    else if (m_currentMode == UNIFIED_QUERY_MODE) {
        restoreUnifiedQuery();
    }
}

ReportDataModel::TemplateType ReportDataModel::getReportType() const {
    return m_reportType;
}

BaseReportParser* ReportDataModel::getParser() const {
    return m_parser;
}

void ReportDataModel::notifyDataChanged() {
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
}

bool ReportDataModel::hasExecutedQueries() const
{
    if (m_reportType != DAY_REPORT && m_reportType != MONTH_REPORT) {
        return false;
    }

    // 检查是否有数据标记被执行过查询
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (cell && cell->cellType == CellData::DataMarker && cell->queryExecuted) {
            return true;
        }
    }

    return false;
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

QVariant ReportDataModel::getTemplateCellData(const QModelIndex& index, int role) const
{
    if (!index.isValid()) return QVariant();
    // ===== 这里放置原 data() 函数中的所有逻辑 =====
    const CellData* cell = getCell(index.row(), index.column());
    if (!cell) return QVariant();

    switch (role) {
    case Qt::DisplayRole:
        return cell->displayText();
    case Qt::EditRole:
        return cell->editText();
    case Qt::BackgroundRole:
        if (cell->cellType == CellData::DataMarker &&
            cell->queryExecuted && !cell->querySuccess) {
            return QBrush(QColor(255, 220, 220));
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

QVariant ReportDataModel::data(const QModelIndex& index, int role) const
{
    if (!index.isValid())
        return QVariant();

    // ===== 根据模式分发 =====
    if (m_currentMode == TEMPLATE_MODE) {
        return getTemplateCellData(index, role);
    }
    else if (m_currentMode == UNIFIED_QUERY_MODE) {
        return getUnifiedQueryCellData(index, role);
    }

    return QVariant();
}


bool ReportDataModel::refreshTemplateReport(QProgressDialog* progress)
{
        // ===== 【步骤1】检测变化类型 =====
        ChangeType changeType = detectChanges();
    
        QString changeMsg;
        bool needQuery = false;
    
        switch (changeType) {
        case NO_CHANGE:
            changeMsg = "数据已是最新，无需刷新。";
            QMessageBox::information(nullptr, "无需刷新", changeMsg);
            saveRefreshSnapshot();
            return true;
    
        case FORMULA_ONLY: {
            int newFormulaCount = 0;
            for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
                if (it.value() && it.value()->hasFormula && !it.value()->formulaCalculated) {
                    newFormulaCount++;
                }
            }
            changeMsg = QString("检测到 %1 个新增公式，是否计算？\n\n注意：不会重新查询数据。")
                .arg(newFormulaCount);
            needQuery = false;
            break;
        }
    
        case BINDING_ONLY: {
            int newMarkerCount = 0;
            if (m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) {
                for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
                    const CellData* cell = it.value();
                    if (cell && cell->cellType == CellData::DataMarker && !cell->queryExecuted) {
                        newMarkerCount++;
                    }
                }
            }
            else {
                QList<QString> newBindings = getNewBindings();
                newMarkerCount = newBindings.size();
            }
    
            changeMsg = QString("检测到 %1 个新增数据标记，是否查询？\n\n将查询新数据并重新计算所有公式。")
                .arg(newMarkerCount);
            needQuery = true;
            break;
        }
    
        case MIXED_CHANGE: {
            if (m_isFirstRefresh) {
                needQuery = true;
            }
            else {
                int newFormulaCount = 0;
                for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
                    if (it.value() && it.value()->hasFormula && !it.value()->formulaCalculated) {
                        newFormulaCount++;
                    }
                }
    
                int newDataCount = 0;
                if (m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) {
                    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
                        const CellData* cell = it.value();
                        if (cell && cell->cellType == CellData::DataMarker && !cell->queryExecuted) {
                            newDataCount++;
                        }
                    }
                }
                else {
                    newDataCount = getNewBindings().size();
                }
    
                changeMsg = QString("检测到数据变更：\n"
                    "• 新增数据标记: %1 个\n"
                    "• 新增公式: %2 个\n\n"
                    "是否执行完整刷新？")
                    .arg(newDataCount)
                    .arg(newFormulaCount);
                needQuery = true;
            }
            break;
        }
        }
    
        // ===== 【步骤2】用户确认（首次刷新跳过） =====
        if (!m_isFirstRefresh && !changeMsg.isEmpty()) {
            auto reply = QMessageBox::question(nullptr, "确认刷新", changeMsg,
                QMessageBox::Yes | QMessageBox::No,
                QMessageBox::Yes);
            if (reply == QMessageBox::No) {
                return false;
            }
        }
    
        if (needQuery && (m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) && m_parser) {
            if (!m_parser->isCacheValid()) {
                qDebug() << "缓存已过期，将重新查询";
                m_parser->invalidateCache();
            }
        }
    
        // ===== 【步骤3】执行刷新操作 =====
        bool completedSuccessfully = false;
        int actualSuccessCount = 0;
    
        // 根据 needQuery 决定是否查询
        if (needQuery) {
            // 需要查询数据
            if ((m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) && m_parser) {
                completedSuccessfully = m_parser->executeQueries(progress);
    
                // 统计实际成功的数据点
                for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
                    const CellData* cell = it.value();
                    if (cell && cell->cellType == CellData::DataMarker &&
                        cell->queryExecuted && cell->querySuccess) {
                        actualSuccessCount++;
                    }
                }
            }
            else if (m_reportType == NORMAL_EXCEL) {
                qDebug() << "普通Excel不支持数据查询";
                completedSuccessfully = true;
            }
        }
        else {
            // 不需要查询，直接标记为成功
            completedSuccessfully = true;
        }
    
        // 在查询完成后添加
        if (completedSuccessfully) {
            qDebug() << "数据填充完成，开始计算公式...";
            recalculateAllFormulas();
    
            // 执行内存优化
            optimizeMemory();
        }
    
        // ===== 【修改】在数据填充完成后，再计算公式 =====
        qDebug() << "数据填充完成，开始计算公式...";
        recalculateAllFormulas();
        
        optimizeMemory();
    
        bool shouldEnterRunMode = false;
    
        // 保存快照
        if ((m_reportType == DAY_REPORT || m_reportType == MONTH_REPORT) && m_parser) {
            int totalTasks = m_parser->getPendingQueryCount();
            if (totalTasks > 0) {
                double successRate = static_cast<double>(actualSuccessCount) / totalTasks;
                if (successRate >= 0.5) {
                    saveRefreshSnapshot();
                    shouldEnterRunMode = true;  // 进入运行模式
                    qDebug() << QString("保存快照: 成功率 %1%").arg(successRate * 100, 0, 'f', 1);
                }
                else {
                    shouldEnterRunMode = false;  // 保持编辑模式，允许用户修改重试
                    qWarning() << QString("成功率过低 (%1%)，不保存快照").arg(successRate * 100, 0, 'f', 1);
                }
            }
            else {
                saveRefreshSnapshot();
                shouldEnterRunMode = true;
            }
        }
        else {
            saveRefreshSnapshot();
            shouldEnterRunMode = true;
        }
    
        if (shouldEnterRunMode) {
            setEditMode(false);
        }
    
        return completedSuccessfully;
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

        m_dirtyFormulas.insert(QPoint(index.row(), index.column()));
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

        if (!cell->hasFormula) {
            markDependentFormulasDirty(index.row(), index.column());
        }
    }

    emit dataChanged(index, index, { role });
    emit cellChanged(index.row(), index.column());
    return true;
}


Qt::ItemFlags ReportDataModel::getTemplateModeFlags(const QModelIndex& index) const
{
    if (!index.isValid()) return Qt::NoItemFlags;

    // ===== 检查预查询状态 =====
    if (m_parser && m_parser->getEditState() == BaseReportParser::PREFETCHING) {
        // 预查询中：只读
        return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
    }

    // 运行模式下只读
    if (!m_editMode) {
        return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
    }

    // 合并单元格判断
    const CellData* cell = getCell(index.row(), index.column());
    if (cell && cell->mergedRange.isMerged()) {
        if (index.row() != cell->mergedRange.startRow ||
            index.column() != cell->mergedRange.startCol) {
            return Qt::ItemIsEnabled;
        }
    }

    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}

Qt::ItemFlags ReportDataModel::getUnifiedQueryModeFlags(const QModelIndex& index) const
{
    if (!index.isValid()) return Qt::NoItemFlags;

    if (hasUnifiedQueryData()) {
        // ===== 报表阶段：时间列和数据列只读，其他列可编辑 =====
        int col = index.column();

        // 第0列是时间列：只读
        if (col == 0) {
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
        }

        // 第1~N列是数据列：只读（N = m_dataColumnCount）
        if (col >= 1 && col <= m_dataColumnCount) {
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
        }

        // 第N+1列及以后：可编辑（用户自定义列，可以写公式）
        // 但合并单元格的非主单元格除外
        const CellData* cell = getCell(index.row(), index.column());
        if (cell && cell->mergedRange.isMerged()) {
            if (index.row() != cell->mergedRange.startRow ||
                index.column() != cell->mergedRange.startCol) {
                return Qt::ItemIsEnabled;
            }
        }

        return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
    }
    else {
        // ===== 配置阶段：仅前2列可编辑 =====
        if (index.column() < 2) {
            return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
        }
        return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
    }
}

Qt::ItemFlags ReportDataModel::flags(const QModelIndex& index) const
{
    if (!index.isValid()) return Qt::NoItemFlags;

    // ===== 根据模式分发 =====
    if (m_currentMode == TEMPLATE_MODE) {
        return getTemplateModeFlags(index);
    }
    else if (m_currentMode == UNIFIED_QUERY_MODE) {
        return getUnifiedQueryModeFlags(index);
    }

    return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
}


// ===== 设置编辑模式 =====
void ReportDataModel::setEditMode(bool editMode)
{
    if (m_editMode == editMode) return;

    m_editMode = editMode;
    emit editModeChanged(editMode);

    // 通知视图更新所有单元格的可编辑状态
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));

    qDebug() << (editMode ? "进入编辑模式" : "进入运行模式");
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


bool ReportDataModel::saveToExcel(const QString& fileName, ExportMode mode)
{
    return ExcelHandler::saveToFile(fileName, this,
        mode == EXPORT_DATA ? ExcelHandler::EXPORT_DATA : ExcelHandler::EXPORT_TEMPLATE);
}


void ReportDataModel::clearAllCells()
{
    if (m_cells.isEmpty() && !m_parser) return;

    beginResetModel();
    qDeleteAll(m_cells);
    m_cells.clear();
    clearSizes();

    delete m_parser;
    m_parser = nullptr;

    m_reportType = NORMAL_EXCEL;
    m_reportName.clear();

    m_dataColumnCount = 0;  

    // 清空快照
    m_lastSnapshot.bindingKeys.clear();
    m_lastSnapshot.formulaCells.clear();
    m_lastSnapshot.dataMarkerCells.clear();
    m_isFirstRefresh = true;

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
    // 不强制扩展到100行26列，使用实际大小
    m_maxRow = newRowCount > 0 ? newRowCount : 100;
    m_maxCol = newColCount > 0 ? newColCount : 26;

    // 确保行高列宽数组大小匹配实际数据
    if (m_rowHeights.size() > m_maxRow) {
        m_rowHeights.resize(m_maxRow);
    }
    if (m_columnWidths.size() > m_maxCol) {
        m_columnWidths.resize(m_maxCol);
    }
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

ReportDataModel::ChangeType ReportDataModel::detectChanges()
{
    // 获取当前状态
    QSet<QString> currentBindings = getCurrentBindings();
    QSet<QPoint> currentFormulas = getCurrentFormulas();

    QSet<QPoint> currentDataMarkers;
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (cell && cell->cellType == CellData::DataMarker) {
            currentDataMarkers.insert(it.key());
        }
    }

    // 如果是首次刷新，返回混合变化（需要完整刷新）
    if (m_lastSnapshot.isEmpty()) {
        qDebug() << "[detectChanges] 首次刷新，执行完整刷新";
        return MIXED_CHANGE;
    }

    // 检测绑定变化
    bool hasNewBindings = false;
    for (const QString& binding : currentBindings) {
        if (!m_lastSnapshot.bindingKeys.contains(binding)) {
            hasNewBindings = true;
            break;
        }
    }

    // 检测 #d# 标记变化
    bool hasNewDataMarkers = false;
    for (const QPoint& pos : currentDataMarkers) {
        if (!m_lastSnapshot.dataMarkerCells.contains(pos)) {
            hasNewDataMarkers = true;
            break;
        }
    }

    // ===== 检测公式变化：任何新增的公式位置都算变化 =====
    bool hasNewFormulas = false;
    int newFormulaCount = 0;
    for (const QPoint& pos : currentFormulas) {
        if (!m_lastSnapshot.formulaCells.contains(pos)) {
            hasNewFormulas = true;
            newFormulaCount++;
        }
    }

    if (hasNewFormulas) {
        qDebug() << QString("[detectChanges] 检测到 %1 个新增公式").arg(newFormulaCount);
    }

    // 判断变化类型
    bool hasDataChange = hasNewBindings || hasNewDataMarkers;

    if (!hasDataChange && !hasNewFormulas) {
        qDebug() << "[detectChanges] 无变化";
        return NO_CHANGE;
    }
    else if (hasNewFormulas && !hasDataChange) {
        qDebug() << "[detectChanges] 仅公式变化";
        return FORMULA_ONLY;
    }
    else if (hasDataChange && !hasNewFormulas) {
        qDebug() << "[detectChanges] 仅数据标记变化";
        return BINDING_ONLY;
    }
    else {
        qDebug() << "[detectChanges] 混合变化";
        return MIXED_CHANGE;
    }
}

void ReportDataModel::saveRefreshSnapshot()
{
    m_lastSnapshot.bindingKeys = getCurrentBindings();
    m_lastSnapshot.formulaCells = getCurrentFormulas();  // ← 这会记录所有公式位置
    m_lastSnapshot.dataMarkerCells.clear();

    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (cell && cell->cellType == CellData::DataMarker) {
            m_lastSnapshot.dataMarkerCells.insert(it.key());
        }
    }

    m_isFirstRefresh = false;

    // ===== 【添加】调试日志 =====
    qDebug() << QString("快照已保存: %1个绑定, %2个公式, %3个数据标记")
        .arg(m_lastSnapshot.bindingKeys.size())
        .arg(m_lastSnapshot.formulaCells.size())
        .arg(m_lastSnapshot.dataMarkerCells.size());
}

QSet<QString> ReportDataModel::getCurrentBindings() const
{
    QSet<QString> bindings;
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (cell && cell->isDataBinding) {
            bindings.insert(cell->bindingKey);
        }
    }
    return bindings;
}

QSet<QPoint> ReportDataModel::getCurrentFormulas() const
{
    QSet<QPoint> formulas;
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        if (it.value() && it.value()->hasFormula) {
            formulas.insert(it.key());
        }
    }
    return formulas;
}

QList<QString> ReportDataModel::getNewBindings() const
{
    QList<QString> newBindings;
    QSet<QString> currentBindings = getCurrentBindings();

    for (const QString& binding : currentBindings) {
        if (!m_lastSnapshot.bindingKeys.contains(binding)) {
            newBindings.append(binding);
        }
    }
    return newBindings;
}

void ReportDataModel::optimizeMemory()
{
    QList<QPoint> emptyKeys;

    // 默认样式参考
    RTCellStyle defaultStyle;

    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();

        if (!cell) {
            emptyKeys.append(it.key());
            continue;
        }

        // 清理条件：
        // 1. 无内容
        // 2. 无公式
        // 3. 无数据绑定
        // 4. 非合并单元格
        // 5. 使用默认样式
        bool isEmpty = cell->value.isNull() || cell->value.toString().isEmpty();
        bool isDefaultStyle =
            cell->style.backgroundColor == defaultStyle.backgroundColor &&
            cell->style.textColor == defaultStyle.textColor &&
            cell->style.alignment == defaultStyle.alignment &&
            cell->style.font == defaultStyle.font;

        if (isEmpty &&
            !cell->hasFormula &&
            !cell->isDataBinding &&
            cell->cellType == CellData::NormalCell &&
            !cell->mergedRange.isMerged() &&
            isDefaultStyle) {
            emptyKeys.append(it.key());
        }
    }

    // 删除空单元格
    for (const QPoint& key : emptyKeys) {
        delete m_cells.take(key);
    }

    if (!emptyKeys.isEmpty()) {
        qDebug() << QString("内存优化：清理 %1 个空单元格").arg(emptyKeys.size());
    }
}

void ReportDataModel::markFormulaDirty(int row, int col)
{
    const CellData* cell = getCell(row, col);
    if (cell && cell->hasFormula) {
        m_dirtyFormulas.insert(QPoint(row, col));
        qDebug() << QString("标记公式为脏: (%1, %2)").arg(row).arg(col);
    }
}

void ReportDataModel::markDependentFormulasDirty(int changedRow, int changedCol)
{
    QString changedAddr = cellAddress(changedRow, changedCol);

    // 遍历所有公式单元格
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (!cell || !cell->hasFormula) continue;

        // 检查公式是否引用了被修改的单元格
        if (cell->formula.contains(changedAddr, Qt::CaseInsensitive)) {
            m_dirtyFormulas.insert(it.key());
        }
    }

    if (!m_dirtyFormulas.isEmpty()) {
        qDebug() << QString("标记 %1 个依赖公式为脏").arg(m_dirtyFormulas.size());
    }
}

bool ReportDataModel::loadUnifiedQueryConfig(const QString& filePath)
{
    // 加载 Excel 到 m_cells，然后调用 Parser

    // 1. 加载 Excel 文件
    if (!loadFromExcelFile(filePath)) {
        return false;
    }

    // 2. 创建统一查询解析器
    m_parser = new UnifiedQueryParser(this, this);

    // 3. 解析配置（从 m_cells 读取）
    if (!m_parser->scanAndParse()) {
        delete m_parser;
        m_parser = nullptr;
        clearAllCells();
        return false;
    }

    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    if (queryParser) {
        const HistoryReportConfig& config = queryParser->getConfig();
        if (config.columns.isEmpty()) {
            qWarning() << "配置为空，未找到有效的列定义";
            delete m_parser;
            m_parser = nullptr;
            clearAllCells();
            return false;
        }
        qDebug() << QString("统一查询配置加载完成：%1 个数据列").arg(config.columns.size());
    }

    qDebug() << "统一查询配置加载完成";
    return true;
}

// ReportDataModel.cpp refreshUnifiedQuery()
bool ReportDataModel::refreshUnifiedQuery(QProgressDialog* progress)
{
    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    if (!queryParser) {
        qWarning() << "解析器类型错误";
        return false;
    }

    // ===== 检测变化类型 =====
    UnifiedQueryChangeType changeType = detectUnifiedQueryChanges();

    if (changeType == UQ_FORMULA_ONLY) {
        // 只有公式变化，只计算公式
        qDebug() << "[统一查询] 仅增量计算公式";
        recalculateAllFormulas();

        QMessageBox::information(nullptr, "公式计算完成",
            "新增公式已计算完成！\n\n"
            "如需重新查询数据，请再次点击刷新。");
        return true;
    }

    if (changeType == UQ_NO_CHANGE && hasUnifiedQueryData()) {
        // 已有数据且无变化，直接返回成功（或者提示用户）
        auto reply = QMessageBox::question(nullptr, "确认刷新",
            "当前数据无变化。\n\n"
            "是否重新查询数据？",
            QMessageBox::Yes | QMessageBox::No,
            QMessageBox::No);

        if (reply == QMessageBox::No) {
            return false;
        }
        // 用户选择 Yes，继续执行查询
    }

    // ===== 执行数据查询 =====
    // 检查时间配置
    if (queryParser->getQueryIntervalSeconds() == 0 &&
        queryParser->getTimeAxis().isEmpty()) {
        qWarning() << "时间配置无效";
        return false;
    }

    // 执行查询
    if (progress) {
        progress->setLabelText("正在查询数据...");
        progress->setRange(0, 0);
    }

    bool success = queryParser->executeQueries(progress);

    if (!success) {
        qWarning() << "查询失败";
        return false;
    }

    // 更新 Model 尺寸
    const QVector<QDateTime>& timeAxis = queryParser->getTimeAxis();
    const HistoryReportConfig& config = queryParser->getConfig();

    if (timeAxis.isEmpty()) {
        qWarning() << "查询未返回数据";
        return false;
    }

    int totalRows = timeAxis.size() + 1;  // +1 表头
    int totalCols = config.columns.size() + 1;  // +1 时间列

    // ===== 保留用户自定义列 =====
    int oldMaxCol = m_maxCol;
    int userDefinedCols = oldMaxCol - m_dataColumnCount - 1;  // 用户添加的列数
    if (userDefinedCols > 0) {
        totalCols += userDefinedCols;
        qDebug() << QString("保留 %1 个用户自定义列").arg(userDefinedCols);
    }

    m_dataColumnCount = config.columns.size();

    beginResetModel();
    updateModelSize(totalRows, totalCols);
    endResetModel();

    // ===== 计算公式 =====
    recalculateAllFormulas();

    if (progress) {
        progress->setValue(100);
    }

    qDebug() << QString("统一查询完成：%1行 × %2列（数据列：%3）")
        .arg(totalRows).arg(totalCols).arg(m_dataColumnCount);

    // 显示成功消息
    QMessageBox::information(nullptr, "查询成功",
        QString("数据查询完成！\n\n"
            "时间点：%1 个\n"
            "数据列：%2 个\n\n"
            "提示：时间列和数据列为只读，您可以在右侧添加自定义列和公式。")
        .arg(timeAxis.size())
        .arg(config.columns.size()));

    return true;
}

QVariant ReportDataModel::getUnifiedQueryCellData(const QModelIndex& index, int role) const
{
    if (!index.isValid()) return QVariant();

    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    if (!queryParser) return QVariant();

    int row = index.row();
    int col = index.column();

    const QVector<QDateTime>& timeAxis = queryParser->getTimeAxis();
    const HistoryReportConfig& config = queryParser->getConfig();
    const QHash<QString, QVector<double>>& data = queryParser->getAlignedData();

    // ===== 判断当前是配置阶段还是报表阶段 =====
    if (timeAxis.isEmpty()) {
        // 【配置阶段】显示真实 cells（2列）
        if (role == Qt::DisplayRole || role == Qt::EditRole) {
            const CellData* cell = getCell(row, col);
            return cell ? cell->value : QVariant();
        }

        if (role == Qt::BackgroundRole) {
            return QBrush(QColor(250, 250, 250));
        }

        if (role == Qt::TextAlignmentRole) {
            return (int)(Qt::AlignVCenter | Qt::AlignLeft);
        }
    }
    else {
        // 【报表阶段】

        // ===== 优先检查是否有用户输入的单元格数据 =====
        const CellData* cell = getCell(row, col);

        // 如果是用户自定义列（col > m_dataColumnCount），优先显示用户数据
        if (col > m_dataColumnCount && cell) {
            if (role == Qt::DisplayRole) {
                if (cell->hasFormula && cell->formulaCalculated) {
                    // 显示公式计算结果
                    return cell->value;
                }
                else if (cell->hasFormula && !cell->formulaCalculated) {
                    // 显示公式文本
                    return cell->formula;
                }
                else {
                    // 显示普通值
                    return cell->value;
                }
            }

            if (role == Qt::EditRole) {
                if (cell->hasFormula) {
                    return cell->formula;  // 编辑时显示公式
                }
                return cell->value;
            }

            if (role == Qt::BackgroundRole) {
                return QBrush(cell->style.backgroundColor);
            }

            if (role == Qt::ForegroundRole) {
                return QBrush(cell->style.textColor);
            }

            if (role == Qt::FontRole) {
                return ensureFontAvailable(cell->style.font);
            }

            if (role == Qt::TextAlignmentRole) {
                return static_cast<int>(cell->style.alignment);
            }
        }

        // ===== 否则显示查询数据（虚拟数据）=====
        if (role == Qt::DisplayRole || role == Qt::EditRole) {
            // 表头行
            if (row == 0) {
                if (col == 0) return "时间";
                if (col - 1 < config.columns.size()) {
                    return config.columns[col - 1].displayName;
                }
                return QVariant();
            }

            // 数据行
            int dataRow = row - 1;
            if (dataRow >= 0 && dataRow < timeAxis.size()) {
                // 时间列
                if (col == 0) {
                    return timeAxis[dataRow].toString("yyyy-MM-dd HH:mm:ss");
                }
                // 数据列
                else if (col - 1 < config.columns.size()) {
                    QString rtuId = config.columns[col - 1].rtuId;
                    if (data.contains(rtuId) && dataRow < data[rtuId].size()) {
                        double value = data[rtuId][dataRow];
                        if (std::isnan(value) || std::isinf(value)) {
                            return "N/A";
                        }
                        return QString::number(value, 'f', 2);
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
                return QBrush(QColor(220, 220, 220));  // 表头灰色
            }

            // ===== 用户自定义列使用不同背景色 =====
            if (col > m_dataColumnCount) {
                return (row % 2 == 0) ? QBrush(QColor(255, 255, 240)) : QBrush(QColor(250, 250, 235));
            }

            return (row % 2 == 0) ? QBrush(Qt::white) : QBrush(QColor(248, 248, 248));
        }

        if (role == Qt::FontRole) {
            QFont font;
            if (row == 0) font.setBold(true);
            return font;
        }
    }

    return QVariant();
}

bool ReportDataModel::hasUnifiedQueryData() const
{
    if (m_currentMode != UNIFIED_QUERY_MODE || !m_parser) {
        return false;
    }

    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    return queryParser && !queryParser->getTimeAxis().isEmpty();
}

void ReportDataModel::setTimeRangeForQuery(const TimeRangeConfig& config)
{
    if (m_currentMode != UNIFIED_QUERY_MODE || !m_parser) {
        qWarning() << "当前不是统一查询模式";
        return;
    }

    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    if (queryParser) {
        queryParser->setTimeRange(config);
        qDebug() << "时间范围已设置";
    }
}

void ReportDataModel::updateEditability()
{
    // 通知视图更新所有单元格的可编辑状态
    emit dataChanged(index(0, 0), index(m_maxRow - 1, m_maxCol - 1));
}

ReportDataModel::UnifiedQueryChangeType ReportDataModel::detectUnifiedQueryChanges()
{
    if (!hasUnifiedQueryData()) {
        // 还没查询过数据，肯定需要查询
        return UQ_NEED_REQUERY;
    }

    // 检查是否有新增公式
    bool hasNewFormulas = false;
    for (int row = 0; row < m_maxRow; ++row) {
        for (int col = m_dataColumnCount + 1; col < m_maxCol; ++col) {
            const CellData* cell = getCell(row, col);
            if (cell && cell->hasFormula && !cell->formulaCalculated) {
                hasNewFormulas = true;
                break;
            }
        }
        if (hasNewFormulas) break;
    }

    if (hasNewFormulas) {
        qDebug() << "[统一查询] 检测到新增公式";
        return UQ_FORMULA_ONLY;
    }

    qDebug() << "[统一查询] 无变化";
    return UQ_NO_CHANGE;
}