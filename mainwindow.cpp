#pragma execution_character_set("utf-8")

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

#include "mainwindow.h"
#include "reportdatamodel.h"
#include "EnhancedTableView.h"
#include "TaosDataFetcher.h"
#include "DayReportParser.h" 


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


void MainWindow::onImportExcel()
{
    QString fileName = QFileDialog::getOpenFileName(this,
        "导入Excel文件", "", "Excel文件 (*.xlsx *.xls)");

    if (fileName.isEmpty()) return;

    m_tableView->clearSpans();

    // =====【已修正】使用统一的加载接口 =====
    if (m_dataModel->loadReportTemplate(fileName)) {
        applyRowColumnSizes();
        m_tableView->updateSpans();

        QFileInfo fileInfo(fileName);
        setWindowTitle(QString("SCADA报表控件 - [%1]").arg(fileInfo.fileName()));

        // 根据加载后的类型决定提示信息
        if (m_dataModel->getReportType() == ReportDataModel::DAY_REPORT) {
            QMessageBox::information(this, "日报已加载",
                "日报模板已加载，请点击 [刷新数据] 按钮填充数据。");
        }
        else {
            QMessageBox::information(this, "成功", "文件导入成功！");
        }
    }
    else {
        QMessageBox::warning(this, "错误", "文件加载或解析失败！\n\n请检查文件格式或模板标记是否正确。");
    }
}

void MainWindow::onExportExcel()
{
    if (m_dataModel->getAllCells().isEmpty()) {
        QMessageBox::information(this, "提示", "当前没有可导出的数据");
        return;
    }

    // 弹出选择对话框
    QMessageBox msgBox;
    msgBox.setWindowTitle("选择导出模式");
    msgBox.setText("请选择导出类型：");
    msgBox.setInformativeText(
        "• 导出数据：公式和标记将被替换为实际值\n"
        "• 导出模板：保留公式和标记，可重新导入");

    QPushButton* dataButton = msgBox.addButton("导出数据", QMessageBox::AcceptRole);
    QPushButton* templateButton = msgBox.addButton("导出模板", QMessageBox::AcceptRole);
    msgBox.addButton("取消", QMessageBox::RejectRole);

    msgBox.exec();

    if (msgBox.clickedButton() == dataButton) {
        exportData();
    }
    else if (msgBox.clickedButton() == templateButton) {
        exportTemplate();
    }
}

void MainWindow::exportData()
{
    if (m_dataModel->isFirstRefresh()) {
        auto reply = QMessageBox::question(this, "确认导出",
            "数据尚未刷新，建议先点击 [刷新数据]。\n是否继续？",
            QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
        if (reply == QMessageBox::No) return;
    }

    QString fileName = QFileDialog::getSaveFileName(this,
        "导出数据", generateFileName("数据"), "Excel文件 (*.xlsx)");
    if (fileName.isEmpty()) return;

    if (m_dataModel->saveToExcel(fileName, ReportDataModel::EXPORT_DATA)) {
        QMessageBox::information(this, "成功", "数据导出成功！");
    }
    else {
        QMessageBox::warning(this, "错误", "数据导出失败！");
    }
}

void MainWindow::exportTemplate()
{
    QString fileName = QFileDialog::getSaveFileName(this,
        "导出模板", generateFileName("模板"), "Excel文件 (*.xlsx)");
    if (fileName.isEmpty()) return;

    if (m_dataModel->saveToExcel(fileName, ReportDataModel::EXPORT_TEMPLATE)) {
        QMessageBox::information(this, "成功", "模板导出成功！");
    }
    else {
        QMessageBox::warning(this, "错误", "模板导出失败！");
    }
}

QString MainWindow::generateFileName(const QString& suffix)
{
    QString name = m_dataModel->getReportName();
    if (name.isEmpty()) name = "报表";
    if (name.startsWith("##")) name = name.mid(2);

    QString time = QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss");
    return QString("%1_%2_%3.xlsx").arg(name).arg(suffix).arg(time);
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

void MainWindow::onRefreshData()
{
    if (m_dataModel->getAllCells().isEmpty()) {
        QMessageBox::information(this, "提示", "当前没有可刷新的数据。\n\n请先通过 [导入] 按钮加载一个报表模板。");
        return; // 直接结束，不执行后续逻辑
    }

    ReportDataModel::ReportType type = m_dataModel->getReportType();

    // 日报模式需要特殊处理，因为它有耗时的查询操作
    if (type == ReportDataModel::DAY_REPORT) {
        DayReportParser* parser = m_dataModel->getDayParser();
        if (!parser || !parser->isValid()) {
            QMessageBox::warning(this, "错误", "日报解析器未初始化或模板无效！");
            return;
        }

        // 1. 创建进度对话框
        // 我们在这里创建它，但它只在需要长时间操作时才会显示
        QProgressDialog progress("正在刷新数据...", "取消", 0, 100, this);
        progress.setWindowModality(Qt::WindowModal);
        progress.setMinimumDuration(500); // 操作超过0.5秒才显示，避免闪烁

        // 2. 如果解析器中存在待查询的任务，就设置进度条的范围并连接信号
        int queryCount = parser->getPendingQueryCount();
        if (queryCount > 0) {
            progress.setRange(0, queryCount);
            connect(parser, &DayReportParser::queryProgress, &progress, &QProgressDialog::setValue);
        }

        // 3. 调用模型的刷新函数，将进度条传入
        // 模型内部会决定是否执行查询，如果执行，解析器就会更新我们这里的进度条
        bool completed = m_dataModel->refreshReportData(&progress);

        // 4. 操作完成后，断开连接，这是一个好习惯
        if (queryCount > 0) {
            disconnect(parser, &DayReportParser::queryProgress, &progress, &QProgressDialog::setValue);
        }

        // 5. 主窗口只负责处理“用户点击取消”的情况
        // 所有其他的提示（如“已是最新”）都由模型自己负责
        if (!completed && progress.wasCanceled()) {
            QMessageBox::warning(this, "已取消", "数据刷新操作已被用户取消。");
            m_dataModel->restoreToTemplate();
        }
    }
    else {
        // 对于普通Excel等其他模式，通常操作很快，不需要进度条
        // 直接调用刷新即可，模型内部会处理相应的弹窗逻辑
        m_dataModel->refreshReportData(nullptr);
    }
}


