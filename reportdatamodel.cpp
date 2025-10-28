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
#include <QPushButton>
#include <QTimer>

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
        m_currentMode = TEMPLATE_MODE;
        qDebug() << "检测到日报模板，开始解析...";

        m_parser = new DayReportParser(this, this);

        // **关键修改**：连接预查询完成信号，但不在这里设置编辑状态
        connect(m_parser, &BaseReportParser::asyncTaskCompleted,
            this, [this](bool success, const QString& message) {
                qDebug() << "日报预查询完成: " << success << message;
                // **预查询完成后恢复编辑模式**
                if (m_parser) {
                    m_parser->setEditState(BaseReportParser::CONFIG_EDIT);
                }
                setEditMode(true);  // 恢复可编辑
            }, Qt::QueuedConnection);


        if (!m_parser->scanAndParse()) {
            clearAllCells();
            return false;
        }

        // **关键修改**：立即设置为可编辑状态
        setEditMode(true);  // 导入后立即可编辑
        notifyDataChanged();
        return true;
    }

    // **关键修改**：连接预查询完成信号
    connect(m_parser, &BaseReportParser::asyncTaskCompleted,
        this, [this](bool success, const QString& message) {
            qDebug() << "月报预查询完成: " << success << message;
            // **预查询完成后恢复编辑模式**
            if (m_parser) {
                m_parser->setEditState(BaseReportParser::CONFIG_EDIT);
            }
            setEditMode(true);  // 恢复可编辑
        }, Qt::QueuedConnection);

    // ===== 识别月报模式 =====
    if (baseName.startsWith("##Month_", Qt::CaseInsensitive)) {
        m_reportType = MONTH_REPORT;
        m_currentMode = TEMPLATE_MODE;
        qDebug() << "检测到月报模板，开始解析...";

        m_parser = new MonthReportParser(this, this);

        if (!m_parser->scanAndParse()) {
            clearAllCells();
            return false;
        }


        // **关键修改**：立即设置为可编辑状态
        setEditMode(true);  // 导入后立即可编辑
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

    int foundUncalculated = 0;
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (cell && cell->hasFormula && !cell->formulaCalculated) {
            // 只有当它不在集合中时才插入，减少冗余
            if (!m_dirtyFormulas.contains(it.key())) {
                m_dirtyFormulas.insert(it.key());
                foundUncalculated++;
            }
        }
    }
    if (foundUncalculated > 0) {
        qDebug() << QString("额外找到了 %1 个未计算的(原始)公式").arg(foundUncalculated);
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
                for (const QPoint& pos : m_dirtyFormulas) {
                    CellData* cell = getCell(pos.x(), pos.y());
                    if (cell) {
                        cell->displayValue = QVariant("#循环引用!"); // 或者 "#ERROR!"
                        cell->formulaCalculated = true; // 标记为已处理 (即使是错误)
                    }
                }
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

void ReportDataModel::resetModelSize(int rows, int cols)
{
    beginResetModel();
    updateModelSize(rows, cols);  // 你自己的逻辑
    endResetModel();
}

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

// 在 reportdatamodel.cpp 文件中

void ReportDataModel::restoreTemplateReport()
{
    qDebug() << "还原模板模式配置...";

    // === 新的还原逻辑：遍历所有单元格 ===
    int restoredMarkers = 0;
    int restoredFormulas = 0;

    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (!cell) continue;

        // 1. 还原标记单元格 (#Date, #t#, #d#) - 检查 markerText
        if (!cell->markerText.isEmpty())
        {
            // 将显示值恢复为原始标记文本
            cell->displayValue = cell->markerText;

            // 如果是数据标记，重置查询状态
            if (cell->cellType == CellData::DataMarker) { // 也可以用 markerText.startsWith("#d#")
                cell->queryExecuted = false;
                cell->querySuccess = false;
            }
            restoredMarkers++;
        }
        // 2. 还原公式单元格 - 检查 hasFormula
        else if (cell->hasFormula)
        {
            // 将显示值恢复为公式文本
            cell->displayValue = cell->formula;
            cell->formulaCalculated = false; // 标记为未计算
            restoredFormulas++;
        }
        // 3. 普通单元格：其 displayValue 在还原时通常保持不变
    }

    qDebug() << QString("模型还原：还原了 %1 个标记单元格, %2 个公式单元格。")
        .arg(restoredMarkers).arg(restoredFormulas);

    // --- 保留后续的清理和状态设置 ---
    m_lastSnapshot.bindingKeys.clear();
    m_lastSnapshot.formulaCells.clear();
    m_lastSnapshot.dataMarkerCells.clear();
    m_isFirstRefresh = true;
    m_dirtyFormulas.clear();
    clearDirtyMarks();
    qDebug() << "脏标记已清空";

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

    if (m_parser && m_parser->isAsyncTaskRunning()) {
        QMessageBox msgBox(QMessageBox::Warning, "操作被阻止",
            "数据查询正在进行中，无法还原配置。\n\n"
            "请等待查询完成或取消查询后再试。",
            QMessageBox::NoButton, nullptr);
        msgBox.setStandardButtons(QMessageBox::Ok);
        msgBox.setButtonText(QMessageBox::Ok, "确定");
        msgBox.exec();
        return;
    }

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
        QMessageBox msgBox(QMessageBox::Warning, "确认还原配置",
            QString("检测到您添加了 %1 个自定义公式/列。\n\n"
                "还原配置将清除所有查询数据和自定义公式。\n\n"
                "是否继续？").arg(formulaCount),
            QMessageBox::NoButton, nullptr); // 注意：父窗口为 nullptr
        QPushButton* yesBtn = msgBox.addButton("是", QMessageBox::YesRole);
        QPushButton* noBtn = msgBox.addButton("否", QMessageBox::NoRole);
        msgBox.setDefaultButton(noBtn); // 保持原始默认值为 No

        msgBox.exec();

        if (msgBox.clickedButton() == noBtn) {
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

        QMessageBox msgBoxDone(QMessageBox::Information, "还原完成", "已还原到配置文件状态。", QMessageBox::NoButton, nullptr);
        msgBoxDone.setStandardButtons(QMessageBox::Ok);
        msgBoxDone.setButtonText(QMessageBox::Ok, "确定");
        msgBoxDone.exec();
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

    const CellData* cell = getCell(index.row(), index.column());
    if (!cell) return QVariant();

    switch (role) {
    case Qt::DisplayRole:
        return cell->displayText();  // 使用封装的方法

    case Qt::EditRole:
        return cell->editText();     // 使用封装的方法

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
    if (!m_parser) {
        qWarning() << "解析器为空";
        return false;
    }

    bool dateMarkerChanged = false;
    // 检查是否有 DateMarker 或 (对月报而言) Date1Marker 在脏单元格中
    for (const QPoint& dirtyPos : m_dirtyCells) {
        const CellData* cell = getCell(dirtyPos.x(), dirtyPos.y());
        if (cell && cell->cellType == CellData::DateMarker) {
            // 对于月报，Date1Marker 被识别为 DateMarker
            dateMarkerChanged = true;
            qWarning() << QString("检测到日期标记变化于 (%1, %2)，将强制重新扫描")
                .arg(dirtyPos.x()).arg(dirtyPos.y());
            break;
        }
    }

    // 如果日期标记被修改，执行强制重新扫描和解析
    if (dateMarkerChanged) {
        qDebug() << "强制执行重新扫描和解析...";
        m_parser->invalidateCache(); // 清空缓存，因为基准日期变了

        // 断开旧的完成信号连接，避免重复触发或状态混乱
        disconnect(m_parser, &BaseReportParser::asyncTaskCompleted, this, nullptr);

        // **重要**：重新连接完成信号（与 loadReportTemplate 中逻辑保持一致）
        connect(m_parser, &BaseReportParser::asyncTaskCompleted,
            this, [this](bool success, const QString& message) {
                qDebug() << "强制重新扫描后的预查询完成: " << success << message;
                if (m_parser) {
                    m_parser->setEditState(BaseReportParser::CONFIG_EDIT);
                }
                // 注意：这里不再需要 setEditMode(true)，因为刷新完成后会根据 fillSuccess 设置
            }, Qt::QueuedConnection);


        if (!m_parser->scanAndParse()) { // 重新扫描并解析，这会更新 m_baseDate 等，并启动新的预查询
            qWarning() << "重新扫描失败";
            QMessageBox::warning(nullptr, "扫描失败", "重新扫描模板标记失败，请检查模板。");
            // 失败后，恢复编辑模式可能比较安全
            setEditMode(true);
            m_parser->setEditState(BaseReportParser::CONFIG_EDIT);
            return false; // 扫描失败则无法继续
        }
        // 重新扫描成功后，认为所有单元格都干净了（因为 parser 内部会处理）
        clearDirtyMarks(); // 清理脏标记
        // 重置刷新状态，使其如同首次刷新一样
        m_isFirstRefresh = true;
        qDebug() << "强制重新扫描完成，将按照首次刷新逻辑继续...";
        // 注意：这里不需要手动调用 refreshTemplateReport，因为当前函数会继续执行后续逻辑
        // 但是需要确保后续逻辑按照首次刷新的路径走
    }

    ChangeType changeType = detectChanges();
    bool hasDirtyCells = !m_dirtyCells.isEmpty();
    bool hasNewFormulas = (changeType == FORMULA_ONLY || changeType == MIXED_CHANGE);

    qDebug() << QString("变化检测: isFirstRefresh=%1, changeType=%2, 脏单元格=%3, 新增公式=%4")
        .arg(m_isFirstRefresh).arg(changeType).arg(m_dirtyCells.size()).arg(hasNewFormulas);

    bool fillSuccess = false; // 保存填充操作的结果
    bool querySuccess = true; // 假设查询成功，除非实际失败

    // ===== 分支 1：首次刷新（或还原后）且无编辑 =====
    if (m_isFirstRefresh && !hasDirtyCells) {
        qDebug() << "首次刷新/还原后刷新 且 无脏单元格，直接从缓存填充";
        fillSuccess = fillDataFromCache(progress);
        if (progress && progress->wasCanceled()) return false;
        if (!fillSuccess) {
            qWarning() << "首次刷新/还原后刷新 时从缓存填充数据失败！";
            // QMessageBox::warning(...); // 错误消息已移到后面统一处理
        }
    }
    // ===== 分支 2：非首次刷新，无变化 =====
    else if (!m_isFirstRefresh && changeType == NO_CHANGE && !hasDirtyCells) {
        qDebug() << "非首次刷新，无变化";
        QMessageBox::information(nullptr, "无需刷新", "数据已是最新，无需刷新。");
        return true; // 无需后续处理，直接返回成功
    }
    // ===== 分支 3：仅公式变化 =====
    else if (!hasDirtyCells && changeType == FORMULA_ONLY) {
        qDebug() << "仅公式变化，只计算公式";
        recalculateAllFormulas();
        saveRefreshSnapshot(); // 仅公式变化也需保存快照
        notifyDataChanged();
        return true; // 公式计算完成，返回成功
    }
    // ===== 分支 4：需要处理数据变化（脏单元格或非首次的绑定变化）=====
    else {
        qDebug() << "需要处理数据变化（脏单元格 或 非首次的绑定变化）";
        bool needQuery = false;
        bool scanNeeded = false;

        // ===== 【增强】差分信息和缓存清理 =====
        BaseReportParser::RescanDiffInfo diffInfo;

        if (hasDirtyCells) {
            qDebug() << QString("检测到 %1 个脏单元格，进行增量扫描").arg(m_dirtyCells.size());

            // 执行增量扫描，获取差分信息
            diffInfo = m_parser->rescanDirtyCells(m_dirtyCells);

            // ===== 【关键判断】是否需要全盘扫描 =====
            if (diffInfo.hasTimeMarkerChange) {
                qDebug() << "检测到时间标记变化，切换为全盘扫描";

                // 清空缓存并重新全盘扫描
                m_parser->invalidateCache();
                m_parser->clearQueryTasks();  // 清空错误任务
                scanNeeded = true;
                needQuery = true;
            }
            else {
                // 只是数据标记变化，使用增量方式

                // 清理受影响的缓存项
                m_parser->cleanupCacheByDiff(diffInfo);

                // 判断是否需要查询
                if (diffInfo.newMarkerCount > 0 || !diffInfo.modifiedMarkers.isEmpty()) {
                    qDebug() << "有新增或修改的标记，需要查询";
                    needQuery = true;
                }
                else if (!diffInfo.removedMarkers.isEmpty()) {
                    qDebug() << "只有删除操作，无需查询";
                    needQuery = false;
                }
                else {
                    qDebug() << "无实质性变化，无需查询";
                    needQuery = false;
                }
            }
        }
        else if (!m_isFirstRefresh && (changeType == BINDING_ONLY || changeType == MIXED_CHANGE)) {
            qDebug() << "检测到绑定变化(非首次刷新)，需要重新扫描并查询";
            m_parser->invalidateCache();
            scanNeeded = true;
            needQuery = true;
        }

        // ===== 执行全盘扫描（如果需要）=====
        if (scanNeeded) {
            qDebug() << "执行重新全盘扫描...";
            if (!m_parser->scanAndParse()) {
                qWarning() << "重新扫描失败";
                QMessageBox::warning(nullptr, "扫描失败", "重新扫描模板标记失败，请检查模板。");
                return false;
            }
        }

        // ===== 统一改为异步查询 =====
        if (needQuery) {
            qDebug() << "启动异步查询任务...";

            // 检查是否有待查询任务
            if (m_parser->getPendingQueryCount() == 0) {
                qDebug() << "没有待查询任务，跳过查询";
            }
            else {
                // 启动异步任务
                m_parser->startAsyncTask();

                // 等待异步任务完成（带进度条）
                if (progress) {
                    progress->setLabelText("正在查询数据库...");
                    progress->setRange(0, 0);  // 不确定进度模式
                }

                // 阻塞等待异步任务完成
                QEventLoop loop;
                bool taskCompleted = false;
                QString taskMessage;

                connect(m_parser, &BaseReportParser::asyncTaskCompleted,
                    &loop, [&](bool success, const QString& message) {
                        taskCompleted = true;
                        querySuccess = success;
                        taskMessage = message;
                        loop.quit();
                    });

                // 定期检查取消
                QTimer cancelCheckTimer;
                connect(&cancelCheckTimer, &QTimer::timeout, [&]() {
                    if (progress && progress->wasCanceled()) {
                        m_parser->requestCancel();
                        loop.quit();
                    }
                    });
                cancelCheckTimer.start(100);  // 每100ms检查一次

                loop.exec();  // 阻塞直到完成
                cancelCheckTimer.stop();

                // 检查是否取消
                if (progress && progress->wasCanceled()) {
                    qDebug() << "用户取消了查询";
                    return false;
                }

                if (taskCompleted) {
                    qDebug() << "异步查询完成：" << taskMessage;
                    if (!querySuccess) {
                        qWarning() << "查询失败/部分失败";
                    }
                }
            }
        }

        // ===== 填充数据 =====
        qDebug() << "开始从缓存填充数据 (fillDataFromCache)...";
        fillSuccess = fillDataFromCache(progress);
        if (progress && progress->wasCanceled()) return false;
    }

    // ===== 后续处理（所有需要填充数据的分支都会执行到这里）=====

    recalculateAllFormulas();
    optimizeMemory();
    clearDirtyMarks();

    // ===== 移动格式化循环到这里！=====
    // 只有在数据填充步骤执行过且成功后才进行格式化
    if (fillSuccess && m_parser) {
        qDebug() << "刷新成功，格式化 Date/Time 标记用于显示..."; // <-- 日志点 A
        int formattedCount = 0;
        for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
            CellData* cell = it.value();
            if (cell && (cell->cellType == CellData::DateMarker || cell->cellType == CellData::TimeMarker)) {
                // <-- 日志点 B (确认循环进入)
                // qDebug() << "Formatting cell (" << it.key().x() << "," << it.key().y() << ")";
                QVariant formattedValue = m_parser->formatDisplayValueForMarker(cell);
                // <-- 日志点 C (确认返回值)
                // qDebug() << "  Formatted value:" << formattedValue;
                if (cell->displayValue != formattedValue) {
                    cell->displayValue = formattedValue;
                    formattedCount++;
                }
            }
        }
        qDebug() << "格式化了" << formattedCount << "个 Date/Time 标记的显示值。"; // <-- 日志点 D
    }
    // ===================================

    // ===== 保存快照并进入运行模式 =====
    if (fillSuccess) {
        saveRefreshSnapshot(); // 保存当前状态（包括格式化后的displayValue）
        setEditMode(false); // 进入运行模式
        qDebug() << "刷新成功完成，进入运行模式";
    }
    else {
        qWarning() << "填充数据失败，保持编辑模式";
        QMessageBox::warning(nullptr, "刷新失败", "从缓存填充数据失败。");
    }

    notifyDataChanged(); // 通知UI更新最终状态
    return fillSuccess;
}

void ReportDataModel::markRowDataMarkersDirty(int row)
{
    for (int col = 0; col < m_maxCol; ++col) {
        const CellData* cell = getCell(row, col);
        if (cell && cell->cellType == CellData::DataMarker) {
            markCellDirty(row, col);
            qDebug() << QString("  同行数据标记受影响: (%1, %2)").arg(row).arg(col);
        }
    }
}

bool ReportDataModel::setData(const QModelIndex& index, const QVariant& value, int role)
{
    if (!index.isValid() || role != Qt::EditRole) {
        return false;
    }

    int row = index.row();
    int col = index.column();

    CellData* cell = ensureCell(row, col);
    if (!cell) {
        return false;
    }

    QString text = value.toString();

    // ===== 快速路径：如果编辑文本没变，不做任何处理 =====
    if (text == cell->editText()) {
        return false;
    }

    // ===== 保存旧状态用于比较 =====
    CellData::CellType oldType = cell->cellType;
    QString oldMarkerText = cell->markerText;
    QString oldRtuId = cell->rtuId;

    // ===== 清理旧状态 =====
    cell->hasFormula = false;
    cell->formulaCalculated = false;

    // ===== 处理公式 =====
    if (text.startsWith("#=#")) {
        cell->hasFormula = true;
        cell->formula = text;
        cell->displayValue = text;
        cell->markerText.clear();
        cell->cellType = CellData::NormalCell;
        cell->formulaCalculated = false;

        m_dirtyFormulas.insert(QPoint(row, col));
    }
    // ===== 处理数据标记 #d# =====
    else if (text.startsWith("#d#", Qt::CaseInsensitive)) {
        QString rtuId = m_parser->extractRtuId(text);

        cell->cellType = CellData::DataMarker;
        cell->markerText = text;
        cell->rtuId = rtuId;
        cell->displayValue = text;
        cell->queryExecuted = false;
        cell->querySuccess = false;

        // **关键修改**：更严格的脏标记判断
        bool needMarkDirty = false;

        // 情况1：新增数据标记
        if (oldType != CellData::DataMarker) {
            needMarkDirty = true;
            qDebug() << QString("新增数据标记: (%1, %2) RTU=%3").arg(row).arg(col).arg(rtuId);
        }
        // 情况2：修改已有数据标记的内容
        else if (oldMarkerText != text || oldRtuId != rtuId) {
            needMarkDirty = true;
            qDebug() << QString("修改数据标记: (%1, %2) %3 -> %4, RTU: %5 -> %6")
                .arg(row).arg(col)
                .arg(oldMarkerText).arg(text)
                .arg(oldRtuId).arg(rtuId);
        }

        if (needMarkDirty) {
            markCellDirty(row, col);
        }
    }
    // ===== 处理时间标记 #t# =====
    else if (text.startsWith("#t#", Qt::CaseInsensitive)) {
        cell->cellType = CellData::TimeMarker;
        cell->markerText = text;
        cell->displayValue = text;

        // **关键修改**：时间标记变化也要标记脏
        bool needMarkDirty = false;

        if (oldType != CellData::TimeMarker) {
            needMarkDirty = true;
            qDebug() << QString("新增时间标记: (%1, %2) %3").arg(row).arg(col).arg(text);
        }
        else if (oldMarkerText != text) {
            needMarkDirty = true;
            qDebug() << QString("修改时间标记: (%1, %2) %3 -> %4")
                .arg(row).arg(col).arg(oldMarkerText).arg(text);
        }

        if (needMarkDirty) {
            markCellDirty(row, col);
            // **重要**：时间标记改变会影响同一行的所有数据标记
            markRowDataMarkersDirty(row);
        }
    }
    // ===== 处理日期标记 #Date =====
    else if (text.startsWith("#Date", Qt::CaseInsensitive)) {
        cell->cellType = CellData::DateMarker;
        cell->markerText = text;
        cell->displayValue = text;

        if (oldType != CellData::DateMarker || oldMarkerText != text) {
            markCellDirty(row, col);
            qDebug() << QString("日期标记变化: (%1, %2) %3").arg(row).arg(col).arg(text);
        }
    }
    // ===== 处理普通文本 =====
    else {
        // **关键修改**：如果之前是标记，现在改成普通文本，标记为脏
        if (oldType != CellData::NormalCell) {
            markCellDirty(row, col);
            qDebug() << QString("删除标记: (%1, %2) 类型=%3").arg(row).arg(col).arg((int)oldType);

            // 如果删除的是时间标记，影响同行数据标记
            if (oldType == CellData::TimeMarker) {
                markRowDataMarkersDirty(row);
            }
        }

        cell->cellType = CellData::NormalCell;
        cell->markerText.clear();
        cell->displayValue = value;
        cell->rtuId.clear();
    }

    // ===== 标记依赖公式为脏 =====
    if (!cell->hasFormula) {
        markDependentFormulasDirty(row, col);
    }

    emit dataChanged(index, index, { role });
    emit cellChanged(row, col);

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
    // 根据模式分发
    if (m_currentMode == UNIFIED_QUERY_MODE) {
        return ExcelHandler::saveUnifiedQueryToFile(fileName, this,
            mode == EXPORT_DATA ? ExcelHandler::EXPORT_DATA : ExcelHandler::EXPORT_TEMPLATE);
    }
    else {
        return ExcelHandler::saveToFile(fileName, this,
            mode == EXPORT_DATA ? ExcelHandler::EXPORT_DATA : ExcelHandler::EXPORT_TEMPLATE);
    }
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
    cell->displayValue = result;
    cell->formulaCalculated = true;  // 标记已计算
}

CellData* ReportDataModel::getCell(int row, int col)
{
    return m_cells.value(QPoint(row, col), nullptr);
}

QVariant ReportDataModel::getCellValueForFormula(int row, int col) const
{
    // ===== 统一查询模式：优先从虚拟数据读取 =====
    if (m_currentMode == UNIFIED_QUERY_MODE) {
        UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
        if (queryParser) {
            const QVector<QDateTime>& timeAxis = queryParser->getTimeAxis();
            const HistoryReportConfig& config = queryParser->getConfig();
            const QHash<QString, QVector<double>>& data = queryParser->getAlignedData();

            // 如果有查询数据
            if (!timeAxis.isEmpty()) {
                // 跳过表头行
                if (row == 0) {
                    return QVariant();
                }

                int dataRow = row - 1;
                if (dataRow >= 0 && dataRow < timeAxis.size()) {
                    // 时间列（第0列）
                    if (col == 0) {
                        return timeAxis[dataRow].toString("yyyy-MM-dd HH:mm:ss");
                    }
                    // 数据列（第1列到第N列）
                    else if (col >= 1 && col <= m_dataColumnCount) {
                        int configIndex = col - 1;
                        if (configIndex < config.columns.size()) {
                            QString rtuId = config.columns[configIndex].rtuId;
                            if (data.contains(rtuId) && dataRow < data[rtuId].size()) {
                                double value = data[rtuId][dataRow];
                                if (std::isnan(value) || std::isinf(value)) {
                                    return QVariant("N/A");  // N/A 视为空值
                                }
                                return value;  // 返回数值，不要转成字符串
                            }
                        }
                    }
                }
            }
        }
    }

    // ===== 模板模式或用户自定义列：从 m_cells 读取 =====
    const CellData* cell = getCell(row, col);
    if (!cell) {
        return QVariant();
    }

    // 如果单元格有公式且已计算，返回计算结果
    if (cell->hasFormula && cell->formulaCalculated) {
        return cell->displayValue;  // 使用 displayValue
    }

    // 如果单元格有公式但未计算，返回 0（避免循环依赖）
    if (cell->hasFormula && !cell->formulaCalculated) {
        qWarning() << QString("引用了未计算的公式单元格: (%1, %2)").arg(row).arg(col);
        return QVariant(0.0);
    }

    // 普通单元格，返回显示值
    return cell->displayValue;  // 使用 displayValue
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

    // 遍历所有数据标记单元格，收集它们的 RTU ID 作为"绑定键"
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (cell && cell->cellType == CellData::DataMarker && !cell->rtuId.isEmpty()) {
            // 使用位置+RTU ID 作为唯一标识（避免不同位置相同RTU被误判为无变化）
            QString bindingKey = QString("%1,%2:%3")
                .arg(it.key().x())
                .arg(it.key().y())
                .arg(cell->rtuId);
            bindings.insert(bindingKey);
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
        // 3. 非数据标记单元格（替代原来的 !isDataBinding）
        // 4. 非合并单元格
        // 5. 使用默认样式
        bool isEmpty = cell->displayValue.isNull() ||
            (cell->displayValue.type() == QVariant::String &&
                cell->displayValue.toString().isEmpty());

        bool isDefaultStyle =
            cell->style.backgroundColor == defaultStyle.backgroundColor &&
            cell->style.textColor == defaultStyle.textColor &&
            cell->style.alignment == defaultStyle.alignment &&
            cell->style.font == defaultStyle.font;

        if (isEmpty &&
            !cell->hasFormula &&
            cell->cellType == CellData::NormalCell &&  // 非数据标记
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
    Q_UNUSED(progress)  // 不再使用同步进度框

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

        QMessageBox msgBox(QMessageBox::Information, "公式计算完成",
            "新增公式已计算完成！\n\n"
            "如需重新查询数据，请再次点击刷新。",
            QMessageBox::NoButton, nullptr);
        msgBox.setStandardButtons(QMessageBox::Ok);
        msgBox.setButtonText(QMessageBox::Ok, "确定");
        msgBox.exec();
        return true;
    }

    if (changeType == UQ_NO_CHANGE && hasUnifiedQueryData()) {
        // 已有数据且无变化，提示用户
        QMessageBox msgBox(QMessageBox::Question, "确认刷新",
            "当前数据无变化。\n\n"
            "是否重新查询数据？",
            QMessageBox::NoButton, nullptr); // 注意：父窗口为 nullptr
        QPushButton* yesBtn = msgBox.addButton("是", QMessageBox::YesRole);
        QPushButton* noBtn = msgBox.addButton("否", QMessageBox::NoRole);
        msgBox.setDefaultButton(noBtn); // 保持原始默认值为 No

        msgBox.exec();

        if (msgBox.clickedButton() == noBtn) {
            return false;
        }
    }

    // ===== 检查时间配置 =====
    if (queryParser->getQueryIntervalSeconds() == 0 &&
        queryParser->getTimeAxis().isEmpty()) {
        qWarning() << "时间配置无效";
        return false;
    }

    // ===== 启动异步查询 =====
    queryParser->startAsyncTask();

    m_isFirstRefresh = false;

    // 立即返回，不等待查询完成
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

    UnifiedQueryParser* queryParser = dynamic_cast<UnifiedQueryParser*>(m_parser);
    if (queryParser) {
        const HistoryReportConfig& currentConfig = queryParser->getConfig();

        // 重新扫描当前配置
        HistoryReportConfig newConfig;
        for (int row = 0; row < m_maxRow; ++row) {
            QModelIndex nameIndex = index(row, 0);
            QModelIndex rtuIndex = index(row, 1);

            QString displayName = data(nameIndex, Qt::DisplayRole).toString().trimmed();
            QString rtuId = data(rtuIndex, Qt::DisplayRole).toString().trimmed();

            if (displayName.isEmpty() && rtuId.isEmpty()) {
                break;
            }

            if (displayName.isEmpty() || rtuId.isEmpty()) {
                continue;
            }

            ReportColumnConfig colConfig;
            colConfig.displayName = displayName;
            colConfig.rtuId = rtuId;
            colConfig.sourceRow = row;
            newConfig.columns.append(colConfig);
        }

        // 对比配置
        if (newConfig.columns.size() != currentConfig.columns.size()) {
            qDebug() << "[统一查询] 配置行数变化："
                << currentConfig.columns.size() << "->" << newConfig.columns.size();
            return UQ_NEED_REQUERY;
        }

        for (int i = 0; i < newConfig.columns.size(); ++i) {
            if (newConfig.columns[i].displayName != currentConfig.columns[i].displayName ||
                newConfig.columns[i].rtuId != currentConfig.columns[i].rtuId) {
                qDebug() << "[统一查询] 配置内容变化：行" << i;
                return UQ_NEED_REQUERY;
            }
        }
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

void ReportDataModel::markAllCellsClean()
{
    m_dirtyCells.clear();
    qDebug() << "所有单元格已标记为干净";
}

bool ReportDataModel::fillDataFromCache(QProgressDialog* progress)
{
    if (!m_parser) {
        return false;
    }

    qDebug() << "开始从缓存填充数据...";

    int successCount = 0;
    int failCount = 0;
    int totalDataMarkers = 0;

    // 统计数据标记总数
    for (auto it = m_cells.constBegin(); it != m_cells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (cell && cell->cellType == CellData::DataMarker) {
            totalDataMarkers++;
        }
    }

    if (progress) {
        progress->setRange(0, totalDataMarkers);
        progress->setLabelText("正在填充数据...");
    }

    int processedCount = 0;

    // 遍历所有数据标记单元格
    for (auto it = m_cells.begin(); it != m_cells.end(); ++it) {
        CellData* cell = it.value();
        if (!cell || cell->cellType != CellData::DataMarker) {
            continue;
        }

        int row = it.key().x();
        int col = it.key().y();

        // ===== 【添加】只为新增行打印日志 =====
        qDebug() << QString("【填充数据】[%1,%2] RTU=%3, cellType=%4")
            .arg(row)
            .arg(col)
            .arg(cell->rtuId)
            .arg((int)cell->cellType);
        // ====================================

        // 根据报表类型构造时间戳
        QDateTime dateTime;

        if (m_reportType == DAY_REPORT) {
            dateTime = constructDateTimeForDayReport(row, col);
        }
        else if (m_reportType == MONTH_REPORT) {
            dateTime = constructDateTimeForMonthReport(row, col);
        }

        // ===== 【修改】新增行打印时间戳 =====
        if (row >= 19 && dateTime.isValid()) {
            int64_t timestamp = dateTime.toMSecsSinceEpoch();
            qDebug() << QString("  → 构造时间戳: %1 (%2)")
                .arg(timestamp)
                .arg(dateTime.toString("yyyy-MM-dd HH:mm:ss"));
        }
        // ===================================

        // 从缓存查找数据
        if (dateTime.isValid()) {
            int64_t timestamp = dateTime.toMSecsSinceEpoch();
            float value = 0.0f;

            if (m_parser->findInCache(cell->rtuId, timestamp, value)) {
                cell->displayValue = QString::number(value, 'f', 2);  // 更新显示值
                cell->queryExecuted = true;
                cell->querySuccess = true;
                successCount++;

                // ===== 【添加】新增行打印结果 =====
                if (row >= 19) {
                    qDebug() << QString("  → 缓存命中: value=%1").arg(value);
                }
                // ==================================
            }
            else {
                cell->displayValue = "N/A";  // 更新显示值
                cell->queryExecuted = true;
                cell->querySuccess = false;
                failCount++;

                // ===== 【添加】新增行打印失败 =====
                if (row >= 19) {
                    qDebug() << QString("  → 缓存未命中");
                }
                // ==================================
            }
        }
        else {
            cell->value = "N/A";
            cell->queryExecuted = true;
            cell->querySuccess = false;
            failCount++;
        }

        processedCount++;
        if (progress) {
            progress->setValue(processedCount);
            if (progress->wasCanceled()) {
                qDebug() << "用户取消填充";
                return false;
            }
        }
    }

    qDebug() << QString("缓存填充完成: 成功 %1, 失败 %2, 总计 %3")
        .arg(successCount).arg(failCount).arg(totalDataMarkers);

    return successCount > 0;
}

// 辅助方法：为日报构造日期时间
QDateTime ReportDataModel::constructDateTimeForDayReport(int row, int col)
{
    DayReportParser* dayParser = dynamic_cast<DayReportParser*>(m_parser);
    if (!dayParser || dayParser->getBaseDate().isEmpty()) {
        qWarning() << "Row" << row << ": DayReportParser invalid or baseDate empty.";
        return QDateTime();
    }

    // 【修复】在同一行向左查找时间标记 (从当前数据单元格的列开始向左)
    // 这保证了我们找到的是逻辑上属于这个数据块的那个#t#标记
    QTime time;
    bool timeFound = false;
    for (int c = col; c >= 0; --c) { // <--- 关键修改：从 col 开始，向左 (c--) 搜索
        const CellData* timeCell = getCell(row, c);

        // ===== 检查 markerText 并从中解析 =====
        if (timeCell && !timeCell->markerText.isEmpty() && timeCell->markerText.startsWith("#t#", Qt::CaseInsensitive)) {
            QString timeMarker = timeCell->markerText;

            // 调用 extractTime 从 markerText 解析时间字符串 "HH:mm:ss"
            // (我们复用 parser 里的解析逻辑，保持一致)
            QString timeStr = m_parser->extractTime(timeMarker);

            if (!timeStr.isEmpty()) {
                time = QTime::fromString(timeStr, "HH:mm:ss");
                if (time.isValid()) {
                    timeFound = true;
                }
                else {
                    qWarning() << "Row" << row << ": Failed to parse extracted time:" << timeStr << "from marker" << timeMarker;
                }
            }
            else {
                qWarning() << "Row" << row << ": extractTime failed for marker:" << timeMarker;
            }

            break; // <--- 关键修改：找到第一个（即最近的）#t#标记后立即停止
        }
    }

    if (timeFound) {
        QString baseDateStr = dayParser->getBaseDate();
        QDate date = QDate::fromString(baseDateStr, "yyyy-MM-dd");
        QDateTime result = QDateTime(date, time);

        return result;
    }
    else {
        // 只有在循环结束后仍未找到时间才打印此警告
        // 【增强调试】加入列信息
        qWarning() << "Row" << row << " Col" << col << ": Could not find valid TimeMarker cell or parse time.";
    }
    return QDateTime();
}

// 辅助方法：为月报构造日期时间
QDateTime ReportDataModel::constructDateTimeForMonthReport(int row, int col)
{
    MonthReportParser* monthParser = dynamic_cast<MonthReportParser*>(m_parser);
    if (!monthParser || monthParser->getBaseYearMonth().isEmpty() || monthParser->getBaseTime().isEmpty()) {
        qWarning() << "Row" << row << ": MonthReportParser invalid or base date/time empty.";
        return QDateTime();
    }

    // 【修复】在同一行查找日期标记 (#t# 代表日)
    // 统一为“从当前数据单元格的列开始向左查找”
    int day = 0;
    bool dayFound = false;
    for (int c = col; c >= 0; --c) { // <--- 关键修改：从 col 开始，向左 (c--) 搜索
        const CellData* dayCell = getCell(row, c);

        // ===== 检查 markerText 并从中解析 =====
        if (dayCell && !dayCell->markerText.isEmpty() && dayCell->markerText.startsWith("#t#", Qt::CaseInsensitive)) {
            QString dayMarker = dayCell->markerText;

            // 直接从 markerText 中提取数字部分
            QString dayStr = dayMarker.mid(3).trimmed();
            bool ok = false;
            day = dayStr.toInt(&ok);

            // 增加对 day 范围的检查
            if (!ok || day < 1 || day > 31) {
                qWarning() << "Row" << row << ": Failed to parse valid day number from markerText:" << dayMarker;
                day = 0; // 重置为无效值
            }
            else {
                dayFound = true;
            }

            break; // <--- 关键修改：找到了标记单元格就跳出内层循环
        }
        // ===========================================
    }

    if (dayFound) {
        QString yearMonth = monthParser->getBaseYearMonth();
        QString baseTimeStr = monthParser->getBaseTime(); // 应为 "HH:mm:ss" 格式
        QString fullDateStr = QString("%1-%2").arg(yearMonth).arg(day, 2, 10, QChar('0'));
        QDate date = QDate::fromString(fullDateStr, "yyyy-MM-dd");
        QTime time = QTime::fromString(baseTimeStr, "HH:mm:ss");

        // 增加对 date 和 time 有效性的检查和日志
        if (date.isValid() && time.isValid()) {
            return QDateTime(date, time);
        }
        else {
            qWarning() << "Row" << row << ": Failed to construct valid QDateTime. Date valid:" << date.isValid() << "Time valid:" << time.isValid() << "DateStr:" << fullDateStr << "TimeStr:" << baseTimeStr;
        }
    }
    else {
        // 只有在循环结束后仍未找到日期才打印此警告
        // 【增强调试】加入列信息
        qWarning() << "Row" << row << " Col" << col << ": Could not find valid TimeMarker cell or parse day number.";
    }

    return QDateTime();
}

void ReportDataModel::markCellDirty(int row, int col)
{
    m_dirtyCells.insert(QPoint(row, col));
}

void ReportDataModel::markRegionDirty(int startRow, int startCol, int endRow, int endCol)
{
    for (int r = startRow; r <= endRow; ++r) {
        for (int c = startCol; c <= endCol; ++c) {
            m_dirtyCells.insert(QPoint(r, c));
        }
    }
}

void ReportDataModel::clearDirtyMarks()
{
    m_dirtyCells.clear();
}