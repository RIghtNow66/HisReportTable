#pragma execution_character_set("utf-8")

#include "mainwindow.h"
#include "reportdatamodel.h"
#include "EnhancedTableView.h"
#include "TaosDataFetcher.h"

#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QHeaderView>
#include <QLineEdit>
#include <QInputDialog>
#include <QProgressDialog>
#include <QSortFilterProxyModel>
#include <QFileInfo>
#include <QDebug>
#include <QShortcut>
#include <QRegularExpression> 


MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
    , m_updating(false)
	, m_formulaEditMode(false)
    , m_filterModel(nullptr)
{

    setupUI();
    setupToolBar();
    setupFormulaBar();
    setupTableView();
    setupContextMenu();
    setupFindDialog();

    setWindowTitle("SCADA报表控件 v1.0");
    resize(1200, 800);
}

MainWindow::~MainWindow()
{
}

void MainWindow::setupUI()
{
    m_centralWidget = new QWidget(this);
    setCentralWidget(m_centralWidget);

    m_mainLayout = new QVBoxLayout(m_centralWidget);
    m_mainLayout->setSpacing(0);
    m_mainLayout->setContentsMargins(0, 0, 0, 0);
}

void MainWindow::setupToolBar()
{
    m_toolBar = addToolBar("主工具栏");

    // 文件操作
    m_toolBar->addAction("导入", this, &MainWindow::onImportExcel);
    m_toolBar->addAction("导出", this, &MainWindow::onExportExcel);
    m_toolBar->addSeparator();

    m_toolBar->addAction("刷新数据", this, &MainWindow::onRefreshData);
    //m_toolBar->addAction("还原配置", this, &MainWindow::onRestoreConfig);
    m_toolBar->addSeparator();

    // 工具操作
    m_toolBar->addAction("查找", this, &MainWindow::onFind);
    m_toolBar->addAction("筛选", this, &MainWindow::onFilter);
    // 清楚筛选
    m_toolBar->addAction("清除筛选", this, &MainWindow::onClearFilter);

    m_toolBar->addSeparator();
}

void MainWindow::setupFormulaBar()
{
    m_formulaWidget = new QWidget();
    QHBoxLayout* layout = new QHBoxLayout(m_formulaWidget);
    layout->setContentsMargins(5, 5, 5, 5);

    m_cellNameLabel = new QLabel("A1");
    m_cellNameLabel->setMinimumWidth(60);
    m_cellNameLabel->setMaximumWidth(60);
    m_cellNameLabel->setAlignment(Qt::AlignCenter);
    m_cellNameLabel->setStyleSheet("QLabel { border: 1px solid gray; padding: 2px; }");

    m_formulaEdit = new QLineEdit();
    m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid gray; padding: 2px; }");

    layout->addWidget(m_cellNameLabel);
    layout->addWidget(m_formulaEdit);

    m_mainLayout->addWidget(m_formulaWidget);

    connect(m_formulaEdit, &QLineEdit::editingFinished, this, &MainWindow::onFormulaEditFinished);
    connect(m_formulaEdit, &QLineEdit::textChanged, this, &MainWindow::onFormulaTextChanged);
}

void MainWindow::setupTableView()
{
    m_tableView = new EnhancedTableView();
    m_dataModel = new ReportDataModel(this);

    m_tableView->setModel(m_dataModel);
    m_tableView->setSelectionBehavior(QAbstractItemView::SelectItems);
    m_tableView->setSelectionMode(QAbstractItemView::SingleSelection);

    m_tableView->setContextMenuPolicy(Qt::CustomContextMenu);

    m_tableView->horizontalHeader()->setDefaultSectionSize(80);
    m_tableView->verticalHeader()->setDefaultSectionSize(25);
    m_tableView->setAlternatingRowColors(false);
    m_tableView->setGridStyle(Qt::SolidLine);

    m_mainLayout->addWidget(m_tableView);

    connect(m_tableView->selectionModel(), &QItemSelectionModel::currentChanged,
        this, &MainWindow::onCurrentCellChanged);
    connect(m_dataModel, &ReportDataModel::cellChanged,
        this, &MainWindow::onCellChanged);
    connect(m_tableView, &QTableView::customContextMenuRequested,
        [this](const QPoint& pos) {
            m_contextMenu->exec(m_tableView->mapToGlobal(pos));
        });

    connect(m_tableView, &QTableView::clicked, this, &MainWindow::onCellClicked);
}

void MainWindow::setupContextMenu()
{
    m_contextMenu = new QMenu(this);

    m_contextMenu->addAction("插入行", this, &MainWindow::onInsertRow);
    m_contextMenu->addAction("插入列", this, &MainWindow::onInsertColumn);
    m_contextMenu->addSeparator();
    m_contextMenu->addAction("删除行", this, &MainWindow::onDeleteRow);
    m_contextMenu->addAction("删除列", this, &MainWindow::onDeleteColumn);
    m_contextMenu->addSeparator();
    m_contextMenu->addAction("向下填充公式", this, &MainWindow::onFillDownFormula);
}

void MainWindow::setupFindDialog()
{
    m_findDialog = new QDialog(this);
    m_findDialog->setWindowTitle("查找");
    m_findDialog->setModal(false);
    m_findDialog->resize(300, 120);

    QVBoxLayout* layout = new QVBoxLayout(m_findDialog);

    QHBoxLayout* inputLayout = new QHBoxLayout();
    inputLayout->addWidget(new QLabel("查找内容:"));
    m_findLineEdit = new QLineEdit();
    inputLayout->addWidget(m_findLineEdit);
    layout->addLayout(inputLayout);

    QHBoxLayout* buttonLayout = new QHBoxLayout();
    QPushButton* findNextBtn = new QPushButton("查找下一个");
    QPushButton* closeBtn = new QPushButton("关闭");
    buttonLayout->addWidget(findNextBtn);
    buttonLayout->addWidget(closeBtn);
    layout->addLayout(buttonLayout);

    connect(findNextBtn, &QPushButton::clicked, this, &MainWindow::onFindNext);
    connect(closeBtn, &QPushButton::clicked, m_findDialog, &QDialog::hide);
    connect(m_findLineEdit, &QLineEdit::returnPressed, this, &MainWindow::onFindNext);
}

void MainWindow::onFillDownFormula()
{
    QModelIndex current = m_tableView->currentIndex();
    if (!current.isValid()) {
        QMessageBox::information(this, "提示", "请先选中一个单元格");
        return;
    }

    int currentRow = current.row();
    int currentCol = current.column();

    // 1. 检查当前单元格是否有公式
    const CellData* sourceCell = m_dataModel->getCell(currentRow, currentCol);
    if (!sourceCell || !sourceCell->hasFormula) {
        QMessageBox::information(this, "提示", "当前单元格没有公式");
        return;
    }

    // 2. 自动判断填充范围
    int endRow = findFillEndRow(currentRow, currentCol);

    if (endRow <= currentRow) {
        QMessageBox::information(this, "提示", "未找到可填充的范围（左侧列没有数据）");
        return;
    }

    // 3. 确认对话框
    auto reply = QMessageBox::question(this, "确认填充",
        QString("是否将公式从第 %1 行填充到第 %2 行？")
        .arg(currentRow + 1)
        .arg(endRow + 1),
        QMessageBox::Yes | QMessageBox::No,
        QMessageBox::Yes);

    if (reply == QMessageBox::No) {
        return;
    }

    // 4. 执行填充
    QString originalFormula = sourceCell->formula;

    for (int row = currentRow + 1; row <= endRow; row++) {
        // 调整公式引用
        QString adjustedFormula = adjustFormulaReferences(originalFormula, row - currentRow);

        QModelIndex targetIndex = m_dataModel->index(row, currentCol);
        m_dataModel->setData(targetIndex, adjustedFormula, Qt::EditRole);
    }

    QMessageBox::information(this, "完成",
        QString("已将公式填充到第 %1 行").arg(endRow + 1));
}

int MainWindow::findFillEndRow(int currentRow, int currentCol)
{
    int maxRow = currentRow;
    int totalCols = m_dataModel->columnCount();

    // 向左查找相邻的列，找到最后一个有数据的行
    for (int col = 0; col < currentCol; col++) {
        // 从当前行开始向下查找，找到该列最后一个非空单元格
        for (int row = currentRow + 1; row < m_dataModel->rowCount(); row++) {
            QModelIndex index = m_dataModel->index(row, col);
            QVariant data = m_dataModel->data(index, Qt::DisplayRole);

            if (!data.isNull() && !data.toString().trimmed().isEmpty()) {
                maxRow = qMax(maxRow, row);
            }
        }
    }

    return maxRow;
}

// 工具函数：调整公式中的行引用
QString MainWindow::adjustFormulaReferences(const QString& formula, int rowOffset)
{
    QString result = formula;

    // 匹配单元格引用：支持 $A$1, $A1, A$1, A1
    QRegularExpression regex(R"((\$?)([A-Z]+)(\$?)(\d+))");
    QRegularExpressionMatchIterator it = regex.globalMatch(formula);

    QList<QRegularExpressionMatch> matches;
    while (it.hasNext()) {
        matches.append(it.next());
    }

    // 从后往前替换
    for (int i = matches.size() - 1; i >= 0; --i) {
        const QRegularExpressionMatch& match = matches[i];
        QString colAbs = match.captured(1);      // $ 或空
        QString colPart = match.captured(2);     // A-Z
        QString rowAbs = match.captured(3);      // $ 或空
        QString rowPart = match.captured(4);     // 数字

        int originalRow = rowPart.toInt();

        // ✅ 只有非绝对引用才调整
        int newRow = rowAbs.isEmpty() ? (originalRow + rowOffset) : originalRow;

        QString newCellRef = colAbs + colPart + rowAbs + QString::number(newRow);

        result.replace(match.capturedStart(), match.capturedLength(), newCellRef);
    }

    return result;
}


void MainWindow::updateTableSpans()
{
    const auto& allCells = m_dataModel->getAllCells();

    // 清除现有的span设置
    // QTableView没有直接的clearAllSpans方法，需要重置

    // 设置新的span
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const QPoint& pos = it.key();
        const CellData* cell = it.value();

        if (cell && cell->mergedRange.isValid() && cell->mergedRange.isMerged()) {
            // 只为主单元格设置span
            if (pos.x() == cell->mergedRange.startRow &&
                pos.y() == cell->mergedRange.startCol) {

                int rowSpan = cell->mergedRange.rowSpan();
                int colSpan = cell->mergedRange.colSpan();

                if (rowSpan > 1 || colSpan > 1) {
                    m_tableView->setSpan(pos.x(), pos.y(), rowSpan, colSpan);
                }
            }
        }
    }
}

void MainWindow::onImportExcel()
{
    QString fileName = QFileDialog::getOpenFileName(this,
        "导入Excel文件", "", "Excel文件 (*.xlsx *.xls)");

    if (fileName.isEmpty()) return;

    // 提取文件名并判断模式
    QFileInfo fileInfo(fileName);
    QString baseFileName = fileInfo.fileName();

    if (baseFileName.startsWith("#REPO_")) {
        // ========== 历史报表模式 ==========
        m_dataModel->setWorkMode(ReportDataModel::HISTORY_MODE);

        // 提取报表名称
        QString reportName = baseFileName.mid(6); // 去掉 "#REPO_"
        reportName = reportName.left(reportName.lastIndexOf('.')); 

        // 加载报表配置
        if (!m_dataModel->loadReportConfig(fileName)) {
            QMessageBox::warning(this, "错误", "配置文件格式错误！\n请检查文件格式是否正确。");
            return;
        }

        // 显示配置文件内容（2列表格）
        m_dataModel->displayConfigFileContent();

        // 更新窗口标题
        setWindowTitle(QString("SCADA报表控件 - [%1]").arg(reportName));

        // 状态栏提示用户下一步操作
        QMessageBox::information(this, "配置已加载",
            QString("报表配置已加载：%1\n\n请点击工具栏的 [刷新数据] 按钮设置时间范围并生成报表。").arg(reportName));

    }
    else {
        // ========== 实时模式（原有逻辑） ==========
        m_dataModel->setWorkMode(ReportDataModel::REALTIME_MODE);

        // 清除旧的合并单元格信息
        m_tableView->clearSpans();

        if (m_dataModel->loadFromExcel(fileName)) {
            updateTableSpans();
            applyRowColumnSizes();
            QMessageBox::information(this, "成功", "文件导入成功！");
        }
        else {
            QMessageBox::warning(this, "错误", "文件导入失败！");
        }

        setWindowTitle("SCADA报表控件 v1.0");
    }
}

void MainWindow::onExportExcel()
{
    ReportDataModel::WorkMode mode = m_dataModel->currentMode();

    if (mode == ReportDataModel::HISTORY_MODE) {
        // ========== 历史模式：导出完整报表 ==========
        if (!m_dataModel->hasHistoryData()) {
            QMessageBox::warning(this, "提示", "当前没有报表数据，请先点击 [刷新数据] 生成报表。");
            return;
        }

        QString fileName = QFileDialog::getSaveFileName(this,
            "导出报表", "", "Excel文件 (*.xlsx)");

        if (fileName.isEmpty()) return;

        QProgressDialog progress("正在导出Excel...", "取消", 0, 100, this);
        progress.setWindowModality(Qt::WindowModal);
        progress.show();

        if (m_dataModel->exportHistoryReportToExcel(fileName, &progress)) {
            QMessageBox::information(this, "成功", "报表已导出！");
        }
        else {
            QMessageBox::warning(this, "失败", "导出失败，请检查文件路径或权限。");
        }

    }
    else {
        // ========== 实时模式：原有逻辑 ==========
        QString fileName = QFileDialog::getSaveFileName(this,
            "导出Excel文件", "", "Excel文件 (*.xlsx)");

        if (!fileName.isEmpty()) {
            if (m_dataModel->saveToExcel(fileName)) {
                QMessageBox::information(this, "成功", "文件导出成功！");
            }
            else {
                QMessageBox::warning(this, "错误", "文件导出失败！");
            }
        }
    }
}

void MainWindow::onFind()
{
    if (m_findDialog->isVisible()) {
        m_findDialog->raise();
        m_findDialog->activateWindow();
    }
    else {
        m_findDialog->show();
    }
    m_findLineEdit->setFocus();
    m_findLineEdit->selectAll();
}

void MainWindow::onFindNext()
{
    QString searchText = m_findLineEdit->text();
    if (searchText.isEmpty()) {
        return;
    }

    // 从当前位置开始查找
    QModelIndex startIndex = m_currentIndex.isValid() ? m_currentIndex : m_dataModel->index(0, 0);
    int startRow = startIndex.row();
    int startCol = startIndex.column();

    // 从下一个单元格开始搜索
    bool found = false;
    for (int row = startRow; row < m_dataModel->rowCount() && !found; ++row) {
        int beginCol = (row == startRow) ? startCol + 1 : 0;
        for (int col = beginCol; col < m_dataModel->columnCount() && !found; ++col) {
            QModelIndex index = m_dataModel->index(row, col);
            QString cellText = m_dataModel->data(index, Qt::DisplayRole).toString();

            if (cellText.contains(searchText, Qt::CaseInsensitive)) {
                m_tableView->setCurrentIndex(index);
                m_tableView->scrollTo(index);
                found = true;
            }
        }
    }

    if (!found) {
        // 从头开始搜索到当前位置
        for (int row = 0; row <= startRow && !found; ++row) {
            int endCol = (row == startRow) ? startCol : m_dataModel->columnCount() - 1;
            for (int col = 0; col <= endCol && !found; ++col) {
                QModelIndex index = m_dataModel->index(row, col);
                QString cellText = m_dataModel->data(index, Qt::DisplayRole).toString();

                if (cellText.contains(searchText, Qt::CaseInsensitive)) {
                    m_tableView->setCurrentIndex(index);
                    m_tableView->scrollTo(index);
                    found = true;
                }
            }
        }
    }

    if (!found) {
        QMessageBox::information(this, "查找", "未找到匹配的内容");
    }
}

void MainWindow::onFilter()
{
    bool ok;
    QString filterText = QInputDialog::getText(this, "筛选",
        "请输入筛选条件（支持通配符*和?）:", QLineEdit::Normal, "", &ok);

    if (ok && !filterText.isEmpty()) {
        if (!m_filterModel) {
            m_filterModel = new QSortFilterProxyModel(this);
            m_filterModel->setSourceModel(m_dataModel);
            m_filterModel->setFilterKeyColumn(-1); // 搜索所有列
            m_tableView->setModel(m_filterModel);
        }

        // 转换通配符为正则表达式
        QString regexPattern = QRegularExpression::escape(filterText);
        regexPattern.replace("\\*", ".*");
        regexPattern.replace("\\?", ".");

        m_filterModel->setFilterRegularExpression(QRegularExpression(regexPattern, QRegularExpression::CaseInsensitiveOption));

        //statusBar()->showMessage(QString("筛选结果：%1 行").arg(m_filterModel->rowCount()), 3000);
    }
}

void MainWindow::onClearFilter()
{
    if (m_filterModel) {
        m_tableView->setModel(m_dataModel);
        delete m_filterModel;
        m_filterModel = nullptr;
        //statusBar()->showMessage("已清除筛选", 2000);
    }
}


void MainWindow::onCurrentCellChanged(const QModelIndex& current, const QModelIndex& previous)
{
    Q_UNUSED(previous)

    if (m_updating)
        return;
	// 如果在公式编辑模式下，不更新当前单元格
    if (m_formulaEditMode)
        return;

    m_currentIndex = current;
    updateFormulaBar(current);
}

void MainWindow::onFormulaEditFinished()
{
    if (m_updating || !m_currentIndex.isValid())
        return;

    // 退出公式编辑模式
    if (m_formulaEditMode)
    {
		exitFormulaEditMode();
    }

    QString text = m_formulaEdit->text();
    m_updating = true;
    m_dataModel->setData(m_currentIndex, text, Qt::EditRole);
    m_updating = false;
}

void MainWindow::enterFormulaEditMode()
{
    m_formulaEditMode = true;
    m_formulaEditingIndex = m_currentIndex;

    // 改变公式编辑框的样式，提示用户进入了公式编辑模式
    m_formulaEdit->setStyleSheet("QLineEdit { border: 2px solid blue; padding: 2px; background-color: #f0f8ff; }");

    // 可以在状态栏显示提示信息
    //statusBar()->showMessage("公式编辑模式：点击单元格插入引用，按Enter完成编辑", 5000);
}

void MainWindow::exitFormulaEditMode()
{
    m_formulaEditMode = false;
    m_formulaEditingIndex = QModelIndex();

    // 恢复公式编辑框的正常样式
    m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid gray; padding: 2px; }");

    //statusBar()->clearMessage();
}

bool MainWindow::isInFormulaEditMode() const
{
    return m_formulaEditMode;
}


void MainWindow::onCellChanged(int row, int col)
{
    if (m_currentIndex.isValid() &&
        m_currentIndex.row() == row &&
        m_currentIndex.column() == col) {
        updateFormulaBar(m_currentIndex);
    }
}

void MainWindow::onCellClicked(const QModelIndex& index)
{
    if (!index.isValid())
    {
        return;
    }
    // 如果在公式编辑状态下，将点击的单元格地址添加到公式中
    if (m_formulaEditMode)
    {
        QString cellAddress = m_dataModel->cellAddress(index.row(), index.column());

        // 获取当前光标位置
        int cursorPos = m_formulaEdit->cursorPosition();
        QString currentText = m_formulaEdit->text();

        // 在光标位置插入单元格地址
        QString newText = currentText.left(cursorPos) + cellAddress + currentText.mid(cursorPos);
        
        m_updating = true;
        m_formulaEdit->setText(newText);
        m_formulaEdit->setCursorPosition(cursorPos + cellAddress.length());
        m_formulaEdit->setFocus(); // 保持焦点在公式编辑框
		m_updating = false;
        
        return;
    }

    // 正常模式下的单元格切换
    if (index != m_currentIndex)
    {
        m_currentIndex = index;
        updateFormulaBar(index);
    }

}

void MainWindow::onFormulaTextChanged()
{
    if (m_updating)
    {
        return;
    }

    QString text = m_formulaEdit->text();

    if (text.startsWith('=') && !m_formulaEditMode)
    {
        enterFormulaEditMode();
    }
    else if (!text.startsWith('=') && m_formulaEditMode)
    {
        exitFormulaEditMode();
    }

}

void MainWindow::updateFormulaBar(const QModelIndex& index)
{
    if (!index.isValid()) {
        m_cellNameLabel->setText("");
        m_formulaEdit->clear();
        return;
    }

    m_updating = true;

	// 更新单元格名称和公式内容
    QString cellName = m_dataModel->cellAddress(index.row(), index.column());
    m_cellNameLabel->setText(cellName);

    //  获取单元格的编辑数据（如果公式则显示公式，否则显示值）
    QVariant editData = m_dataModel->data(index, Qt::EditRole);
    m_formulaEdit->setText(editData.toString());

    // 检查是否应该进入公式编辑模式
    if (editData.toString().startsWith('=')) {
        // 如果新选中的单元格包含公式，但我们不在编辑模式，不要自动进入
        // 只有用户开始编辑时才进入公式编辑模式

        // 可以给公式编辑框一个特殊的样式提示，表明这是一个公式单元格
        m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid #4CAF50; padding: 2px; background-color: #f9fff9; }");
    }
    else {
        // 普通单元格的样式
        m_formulaEdit->setStyleSheet("QLineEdit { border: 1px solid gray; padding: 2px; }");
    }

    m_updating = false;
}

void MainWindow::onInsertRow()
{
    if (m_currentIndex.isValid()) {
        int row = m_currentIndex.row();
        m_dataModel->insertRows(row, 1);
    }
}

void MainWindow::onInsertColumn()
{
    if (m_currentIndex.isValid()) {
        int col = m_currentIndex.column();
        m_dataModel->insertColumns(col, 1);
    }
}

void MainWindow::onDeleteRow()
{
    if (m_currentIndex.isValid()) {
        int row = m_currentIndex.row();
        m_dataModel->removeRows(row, 1);
    }
}

void MainWindow::onDeleteColumn()
{
    if (m_currentIndex.isValid()) {
        int col = m_currentIndex.column();
        m_dataModel->removeColumns(col, 1);
    }
}

void MainWindow::applyRowColumnSizes()
{
    const auto& colWidths = m_dataModel->getAllColumnWidths();
    for (int i = 0; i < colWidths.size(); ++i) {
        if (colWidths[i] > 0) {
            m_tableView->setColumnWidth(i, static_cast<int>(colWidths[i]));
        }
    }

    const auto& rowHeights = m_dataModel->getAllRowHeights();
    for (int i = 0; i < rowHeights.size(); ++i) {
        if (rowHeights[i] > 0) {
            m_tableView->setRowHeight(i, static_cast<int>(rowHeights[i]));
        }
    }
}

// 新增：刷新数据槽函数（预留）
void MainWindow::onRefreshData()
{
    ReportDataModel::WorkMode mode = m_dataModel->currentMode();

    if (mode == ReportDataModel::REALTIME_MODE) {
        // ========== 实时模式：刷新##绑定数据 ==========

        //  检查是否有##绑定
        if (!m_dataModel->hasDataBindings()) {
            QMessageBox::information(this, "提示",
                "当前表格中没有数据绑定（##标记的单元格）。\n\n"
                "如需使用数据绑定功能，请在单元格中输入 ##RTU号 格式。");
            return;
        }

        m_dataModel->resolveDataBindings();
        QMessageBox::information(this, "完成", "实时数据已更新！");

    }
    else if (mode == ReportDataModel::HISTORY_MODE) {
        // ========== 历史模式：弹出时间对话框 ==========

        TimeSettingsDialog dlg(this);

        // 设置对话框标题
        QString reportName = m_dataModel->getReportName();
        dlg.setWindowTitle(QString("设置 [%1] 的时间范围").arg(reportName));

        // 如果之前已经设置过时间，加载上次的值
        if (m_globalConfig.globalTimeRange.startTime.isValid()) {
            dlg.setStartTime(m_globalConfig.globalTimeRange.startTime);
            // 可以继续设置其他参数...
        }
        else {
            // 默认值：今日00:00到现在
            QDateTime now = QDateTime::currentDateTime();
            dlg.setStartTime(QDateTime(now.date(), QTime(0, 0, 0)));
        }

        if (dlg.exec() == QDialog::Accepted) {
            // 更新全局时间配置
            m_globalConfig.globalTimeRange.startTime = dlg.getStartTime();
            m_globalConfig.globalTimeRange.endTime = dlg.getEndTime();
            m_globalConfig.globalTimeRange.intervalSeconds = dlg.getIntervalSeconds();

            // 执行查询并生成报表
            refreshHistoryReport();
        }
        // 用户点取消 → 不做任何操作
    }
}

void MainWindow::refreshHistoryReport()
{
    HistoryReportConfig config = m_dataModel->getHistoryConfig();
    const TimeRangeConfig& timeRange = m_globalConfig.globalTimeRange;

    // 0. 过滤掉空配置行
    QVector<ReportColumnConfig> validColumns;
    for (const auto& col : config.columns) {
        if (!col.displayName.trimmed().isEmpty() && !col.rtuId.trimmed().isEmpty()) {
            validColumns.append(col);
        }
    }

    if (validColumns.isEmpty()) {
        QMessageBox::warning(this, "错误", "配置中没有有效的列（名称和RTU号不能为空）");
        return;
    }

    config.columns = validColumns;

    // 1. 生成时间轴
    QVector<QDateTime> timeAxis = ReportDataModel::generateTimeAxis(timeRange);
    int totalPoints = timeAxis.size();
    int totalColumns = config.columns.size();

    if (totalPoints == 0) {
        QMessageBox::warning(this, "错误", "时间范围配置错误，无法生成时间轴。");
        return;
    }

    // ✅ 新增：单点查询不需要数据量预警
    bool isSinglePoint = (timeRange.intervalSeconds == 0);

    if (!isSinglePoint) {
        // 2. 数据量预警（仅对范围查询）
        if (totalPoints > 50000) {
            auto reply = QMessageBox::question(this, "数据量警告",
                QString("将生成 %1 行数据（约 %2 MB内存），\n"
                    "查询和渲染可能需要较长时间。\n\n"
                    "建议：\n"
                    "• 增大时间间隔（当前：%3秒）\n"
                    "• 缩短时间范围\n\n"
                    "是否继续？")
                .arg(totalPoints)
                .arg(totalPoints * totalColumns * 8 / 1024 / 1024)
                .arg(timeRange.intervalSeconds),
                QMessageBox::Yes | QMessageBox::No,
                QMessageBox::No);

            if (reply == QMessageBox::No) return;
        }
        else if (totalPoints > 20000) {
            QMessageBox::information(this, "提示",
                QString("将生成 %1 行数据，可能需要 10-30 秒，请耐心等待...").arg(totalPoints));
        }
    }

    // 3. 显示进度条
    QString progressText = isSinglePoint ? "正在查询单点数据..." : "正在查询历史数据...";
    QProgressDialog progress(progressText, "取消", 0, 100, this);
    progress.setWindowModality(Qt::WindowModal);
    progress.setMinimumDuration(0);
    progress.show();
    qApp->processEvents();

    // 4. 批量查询所有RTU
    TaosDataFetcher fetcher;
    QHash<QString, std::map<int64_t, std::vector<float>>> rawData;
    QStringList failedRTUs;

    for (int i = 0; i < totalColumns; i++) {
        if (progress.wasCanceled()) {
            QMessageBox::information(this, "已取消", "数据查询已取消。");
            return;
        }

        const ReportColumnConfig& col = config.columns[i];

        // ✅ 单点查询：查询用户指定的时间点
        QString address;
        if (isSinglePoint) {
            // 单点查询：起始时间为用户指定时间，结束时间为起始时间+1分钟，间隔1秒
            QDateTime endTime = timeRange.startTime.addSecs(60);
            address = QString("%1@%2~%3#1")
                .arg(col.rtuId)
                .arg(timeRange.startTime.toString("yyyy-MM-dd HH:mm:ss"))
                .arg(endTime.toString("yyyy-MM-dd HH:mm:ss"));
        }
        else {
            // 范围查询：正常构造地址
            address = QString("%1@%2~%3#%4")
                .arg(col.rtuId)
                .arg(timeRange.startTime.toString("yyyy-MM-dd HH:mm:ss"))
                .arg(timeRange.endTime.toString("yyyy-MM-dd HH:mm:ss"))
                .arg(timeRange.intervalSeconds);
        }

        try {
            auto data = fetcher.fetchDataFromAddress(address.toStdString());

            if (data.empty()) {
                qWarning() << "RTU无数据:" << col.rtuId;
                failedRTUs.append(QString("%1 (无数据)").arg(col.displayName));
                rawData[col.rtuId] = {};
            }
            else {
                rawData[col.rtuId] = data;
            }

        }
        catch (const std::exception& e) {
            qWarning() << "RTU查询失败:" << col.rtuId << e.what();
            failedRTUs.append(QString("%1 (查询失败: %2)").arg(col.displayName).arg(e.what()));
            rawData[col.rtuId] = {};
        }

        progress.setValue(10 + i * 70 / totalColumns);
        qApp->processEvents();
    }

    // 5. 数据对齐（单点查询不需要插值，直接使用查询结果的第一个点）
    QHash<QString, QVector<double>> alignedData;

    if (isSinglePoint) {
        // ✅ 单点查询：直接提取第一个时间点的数据，不做插值
        for (auto it = rawData.constBegin(); it != rawData.constEnd(); ++it) {
            QString rtuId = it.key();
            const std::map<int64_t, std::vector<float>>& dataMap = it.value();

            QVector<double> singleValue;
            if (!dataMap.empty()) {
                // 取第一个数据点（就是用户查询的时间点）
                float value = dataMap.begin()->second[0];
                singleValue.append(static_cast<double>(value));
            }
            else {
                // 无数据时填充NaN
                singleValue.append(std::numeric_limits<double>::quiet_NaN());
            }
            alignedData[rtuId] = singleValue;
        }
    }
    else {
        // 范围查询：使用线性插值对齐
        alignedData = ReportDataModel::alignDataWithInterpolation(rawData, timeAxis);
    }

    progress.setValue(90);

    // 6. 生成表格
    m_dataModel->generateHistoryReport(config, alignedData, timeAxis);
    progress.setValue(100);

    // 7. 完成提示
    QString message;
    if (isSinglePoint) {
        message = QString("单点查询完成：%1 列\n时间点：%2")
            .arg(totalColumns + 1)
            .arg(timeRange.startTime.toString("yyyy-MM-dd HH:mm:ss"));
    }
    else {
        message = QString("报表生成完成：%1 行 × %2 列")
            .arg(totalPoints + 1)
            .arg(totalColumns + 1);
    }

    if (!failedRTUs.isEmpty()) {
        message += QString("\n\n以下列查询失败或无数据：\n%1")
            .arg(failedRTUs.join("\n"));
        QMessageBox::warning(this, "部分数据缺失", message);
    }
    else {
        QMessageBox::information(this, "成功", message);
    }
}


