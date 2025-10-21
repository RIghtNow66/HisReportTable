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
    m_toolBar->addAction("还原配置", this, &MainWindow::onRestoreConfig);
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

        //  只有非绝对引用才调整
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
    m_tableView->resetColumnWidthsBase();

    // =====【修改】在加载前就建立信号连接 =====
    if (m_dataModel->loadReportTemplate(fileName)) {
        applyRowColumnSizes();
        m_tableView->updateSpans();

        QFileInfo fileInfo(fileName);
        setWindowTitle(QString("SCADA报表控件 - [%1]").arg(fileInfo.fileName()));

        ReportDataModel::ReportType reportType = m_dataModel->getReportType();


        // ===== 【关键修改】对于日报和月报，立即建立信号连接 =====
        if (reportType == ReportDataModel::DAY_REPORT ||
            reportType == ReportDataModel::MONTH_REPORT) {

            BaseReportParser* parser = m_dataModel->getParser();
            if (parser) 

                // 先断开旧连接
                disconnect(parser, &BaseReportParser::prefetchCompleted, this, nullptr);

                // 建立新连接
                connect(parser, &BaseReportParser::prefetchCompleted,
                    this, [this](bool hasData, int dataCount, int successCount, int totalCount) {

                        if (!hasData) {
                            QMessageBox::critical(this, "预查询失败",
                                "数据加载失败，未缓存任何数据！\n\n请检查数据库连接或稍后重试。");
                        }
                        else if (totalCount == 0 || successCount == totalCount) {
                            QMessageBox::information(this, "预查询完成",
                                QString("数据加载完成！\n\n已缓存 %1 个数据点\n\n现在可以点击 [刷新数据] 填充报表。")
                                .arg(dataCount));
                        }
                        else if (successCount > 0) {
                            int failCount = totalCount - successCount;
                            QMessageBox::warning(this, "预查询部分完成",
                                QString("数据加载部分完成！\n\n"
                                    "成功：%1/%2 次查询\n"
                                    "失败：%3 次查询\n"
                                    "已缓存：%4 个数据点\n\n"
                                    "部分数据可能显示为 N/A，建议检查数据库连接。")
                                .arg(successCount).arg(totalCount).arg(failCount).arg(dataCount));
                        }
                    }, Qt::QueuedConnection);

            }

            // ===== 月报和日报都不立即弹窗，等预查询完成 =====        }
        else {
            // 普通Excel立即弹窗
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
    if (!m_currentIndex.isValid()) {
        QMessageBox::information(this, "提示", "请先选中一个单元格");
        return;
    }

    int row = m_currentIndex.row();
    int insertRow = 0;
    int count = 0;

    if (showInsertRowDialog(row, insertRow, count)) {
        qDebug() << QString("插入 %1 行在第 %2 行位置").arg(count).arg(insertRow + 1);
        m_dataModel->insertRows(insertRow, count);
    }
}

void MainWindow::onInsertColumn()
{
    if (!m_currentIndex.isValid()) {
        QMessageBox::information(this, "提示", "请先选中一个单元格");
        return;
    }

    int col = m_currentIndex.column();
    int insertCol = 0;
    int count = 0;

    if (showInsertColumnDialog(col, insertCol, count)) {
        qDebug() << QString("插入 %1 列在第 %2 列位置").arg(count).arg(insertCol + 1);
        m_dataModel->insertColumns(insertCol, count);
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
	// 保存初始行高，供缩放使用
    m_tableView->saveBaseRowHeights();
}

void MainWindow::onRefreshData()
{
    if (m_dataModel->getAllCells().isEmpty()) {
        QMessageBox::information(this, "提示", "当前没有可刷新的数据。\n\n请先通过 [导入] 按钮加载一个报表模板。");
        return;
    }

    ReportDataModel::ReportType type = m_dataModel->getReportType();

    if (type == ReportDataModel::DAY_REPORT || type == ReportDataModel::MONTH_REPORT) {
        BaseReportParser* parser = m_dataModel->getParser();
        if (!parser || !parser->isValid()) {
            QMessageBox::warning(this, "错误", "报表解析器未初始化或模板无效！");
            return;
        }

        if (parser->isPrefetching()) {
            QMessageBox::information(this, "请稍候",
                "数据正在后台加载中，请等待预查询完成后再刷新。\n\n"
                "预查询完成后会自动弹窗提示。");
            return;
        }

        // 1. 创建进度对话框
        QProgressDialog progress("正在准备数据...", "取消", 0, 0, this);
        progress.setWindowModality(Qt::WindowModal);
        progress.setMinimumDuration(500);

        bool needsPrefetch = parser->getPendingQueryCount() > 0 &&
            !parser->isPrefetching();

        if (needsPrefetch) {
            progress.setLabelText("正在预查询数据...");
            progress.setRange(0, 0);

            connect(parser, &BaseReportParser::prefetchProgress,
                &progress, [&progress](int current, int total) {
                    if (progress.maximum() != total) {
                        progress.setRange(0, total);
                    }
                    progress.setValue(current);
                });
        }
        else {
            progress.setLabelText("正在填充数据...");
            int queryCount = parser->getPendingQueryCount();
            if (queryCount > 0) {
                progress.setRange(0, queryCount);
                connect(parser, &BaseReportParser::queryProgress,
                    &progress, &QProgressDialog::setValue);
            }
        }

        bool completed = m_dataModel->refreshReportData(&progress);

        disconnect(parser, nullptr, &progress, nullptr);

        if (!completed && progress.wasCanceled()) {
            QMessageBox::warning(this, "已取消", "数据刷新操作已被用户取消。");
            m_dataModel->restoreToTemplate();
        }
    }
    else {
        m_dataModel->refreshReportData(nullptr);
    }
}

void MainWindow::onRestoreConfig()
{
    if (m_dataModel->rowCount() == 0 || m_dataModel->columnCount() == 0) {
        QMessageBox::information(this, "提示", "当前没有可还原的配置。");
        return;
    }

    if (m_dataModel->isFirstRefresh() && !m_dataModel->hasExecutedQueries()) {
        QMessageBox::information(this, "提示", "当前已经是配置文件状态，无需还原。");
        return;
    }

    // 弹出确认对话框
    QMessageBox::StandardButton reply;
    reply = QMessageBox::question(this, "确认", "是否要将报表还原到模板的初始状态？\n\n所有已填充的数据和计算结果都将被清除。",
        QMessageBox::Yes | QMessageBox::No);

    if (reply == QMessageBox::Yes) {
        // 调用数据模型的还原函数
        m_dataModel->restoreToTemplate();
        QMessageBox::information(this, "完成", "配置已成功还原。");
    }
}

MainWindow::MergeConflictInfo MainWindow::checkRowInsertConflict(int row)
{
    MergeConflictInfo info;
    info.hasConflict = false;
    info.safePosition = row;

    const auto& allCells = m_dataModel->getAllCells();

    // 检查当前行和上下行是否存在跨越此行的合并单元格
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (!cell || !cell->mergedRange.isMerged()) continue;

        const RTMergedRange& range = cell->mergedRange;

        // 情况1：合并区域跨越插入位置（纵向合并）
        if (range.startRow < row && range.endRow >= row) {
            info.hasConflict = true;
            info.message = QString(
                "当前位置存在纵向合并单元格（第%1-%2行）。\n"
                "插入可能破坏合并区域的完整性。\n\n"
                "建议：在第%3行上方或第%4行下方插入。"
            ).arg(range.startRow + 1)
                .arg(range.endRow + 1)
                .arg(range.startRow + 1)
                .arg(range.endRow + 1);

            info.safePosition = range.endRow + 1;
            break;
        }

        // 情况2：正好在合并区域的起始行
        if (range.startRow == row && range.rowSpan() > 1) {
            info.hasConflict = true;
            info.message = QString(
                "第%1行是合并单元格的起始行（合并至第%2行）。\n"
                "建议在下方（第%3行之后）插入以保持合并区域完整。"
            ).arg(row + 1)
                .arg(range.endRow + 1)
                .arg(range.endRow + 1);
            break;
        }
    }

    return info;
}

MainWindow::MergeConflictInfo MainWindow::checkColumnInsertConflict(int col)
{
    MergeConflictInfo info;
    info.hasConflict = false;
    info.safePosition = col;

    const auto& allCells = m_dataModel->getAllCells();

    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const CellData* cell = it.value();
        if (!cell || !cell->mergedRange.isMerged()) continue;

        const RTMergedRange& range = cell->mergedRange;

        // 情况1：合并区域跨越插入位置（横向合并）
        if (range.startCol < col && range.endCol >= col) {
            QString startCol, endCol;
            int temp = range.startCol;
            while (temp >= 0) {
                startCol.prepend(QChar('A' + (temp % 26)));
                temp = temp / 26 - 1;
            }
            temp = range.endCol;
            while (temp >= 0) {
                endCol.prepend(QChar('A' + (temp % 26)));
                temp = temp / 26 - 1;
            }

            info.hasConflict = true;
            info.message = QString(
                "当前位置存在横向合并单元格（%1-%2列）。\n"
                "插入可能破坏合并区域的完整性。\n\n"
                "建议：在%3列左侧或%4列右侧插入。"
            ).arg(startCol).arg(endCol).arg(startCol).arg(endCol);

            info.safePosition = range.endCol + 1;
            break;
        }

        // 情况2：正好在合并区域的起始列
        if (range.startCol == col && range.colSpan() > 1) {
            QString colName;
            int temp = col;
            while (temp >= 0) {
                colName.prepend(QChar('A' + (temp % 26)));
                temp = temp / 26 - 1;
            }

            info.hasConflict = true;
            info.message = QString(
                "%1列是合并单元格的起始列。\n"
                "建议在右侧插入以保持合并区域完整。"
            ).arg(colName);
            break;
        }
    }

    return info;
}

bool MainWindow::showInsertRowDialog(int currentRow, int& insertRow, int& count)
{
    // 检测冲突
    MergeConflictInfo conflict = checkRowInsertConflict(currentRow);

    // 创建对话框
    QDialog dialog(this);
    dialog.setWindowTitle("插入行");
    dialog.setModal(true);
    dialog.resize(380, conflict.hasConflict ? 280 : 200);

    QVBoxLayout* mainLayout = new QVBoxLayout(&dialog);

    // 位置选择组
    QGroupBox* positionGroup = new QGroupBox("插入位置");
    QVBoxLayout* positionLayout = new QVBoxLayout(positionGroup);

    QRadioButton* beforeRadio = new QRadioButton(
        QString("在第 %1 行上方插入").arg(currentRow + 1));
    QRadioButton* afterRadio = new QRadioButton(
        QString("在第 %1 行下方插入").arg(currentRow + 1));

    beforeRadio->setChecked(true);
    positionLayout->addWidget(beforeRadio);
    positionLayout->addWidget(afterRadio);
    mainLayout->addWidget(positionGroup);

    // 数量输入
    QHBoxLayout* countLayout = new QHBoxLayout();
    countLayout->addWidget(new QLabel("插入数量:"));
    QSpinBox* countSpinBox = new QSpinBox();
    countSpinBox->setRange(1, 100);
    countSpinBox->setValue(1);
    countSpinBox->setSuffix(" 行");
    countLayout->addWidget(countSpinBox);
    countLayout->addStretch();
    mainLayout->addLayout(countLayout);

    // 冲突警告
    if (conflict.hasConflict) {
        QLabel* warningLabel = new QLabel();
        warningLabel->setText("⚠️ " + conflict.message);
        warningLabel->setStyleSheet(
            "QLabel { "
            "color: #d32f2f; "
            "background-color: #ffebee; "
            "border: 1px solid #ef5350; "
            "border-radius: 4px; "
            "padding: 8px; "
            "}"
        );
        warningLabel->setWordWrap(true);
        mainLayout->addWidget(warningLabel);
    }

    // 按钮
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    buttonLayout->addStretch();
    QPushButton* okButton = new QPushButton("确定");
    QPushButton* cancelButton = new QPushButton("取消");

    okButton->setDefault(true);
    okButton->setMinimumWidth(80);
    cancelButton->setMinimumWidth(80);

    buttonLayout->addWidget(okButton);
    buttonLayout->addWidget(cancelButton);
    mainLayout->addLayout(buttonLayout);

    connect(okButton, &QPushButton::clicked, &dialog, &QDialog::accept);
    connect(cancelButton, &QPushButton::clicked, &dialog, &QDialog::reject);

    // 显示对话框
    if (dialog.exec() == QDialog::Accepted) {
        count = countSpinBox->value();
        insertRow = beforeRadio->isChecked() ? currentRow : currentRow + 1;
        return true;
    }

    return false;
}

bool MainWindow::showInsertColumnDialog(int currentCol, int& insertCol, int& count)
{
    // 检测冲突
    MergeConflictInfo conflict = checkColumnInsertConflict(currentCol);

    // 获取列名
    QString colName;
    int temp = currentCol;
    while (temp >= 0) {
        colName.prepend(QChar('A' + (temp % 26)));
        temp = temp / 26 - 1;
    }

    // 创建对话框
    QDialog dialog(this);
    dialog.setWindowTitle("插入列");
    dialog.setModal(true);
    dialog.resize(380, conflict.hasConflict ? 280 : 200);

    QVBoxLayout* mainLayout = new QVBoxLayout(&dialog);

    // 位置选择组
    QGroupBox* positionGroup = new QGroupBox("插入位置");
    QVBoxLayout* positionLayout = new QVBoxLayout(positionGroup);

    QRadioButton* beforeRadio = new QRadioButton(
        QString("在 %1 列左侧插入").arg(colName));
    QRadioButton* afterRadio = new QRadioButton(
        QString("在 %1 列右侧插入").arg(colName));

    beforeRadio->setChecked(true);
    positionLayout->addWidget(beforeRadio);
    positionLayout->addWidget(afterRadio);
    mainLayout->addWidget(positionGroup);

    // 数量输入
    QHBoxLayout* countLayout = new QHBoxLayout();
    countLayout->addWidget(new QLabel("插入数量:"));
    QSpinBox* countSpinBox = new QSpinBox();
    countSpinBox->setRange(1, 100);
    countSpinBox->setValue(1);
    countSpinBox->setSuffix(" 列");
    countLayout->addWidget(countSpinBox);
    countLayout->addStretch();
    mainLayout->addLayout(countLayout);

    // 冲突警告
    if (conflict.hasConflict) {
        QLabel* warningLabel = new QLabel();
        warningLabel->setText("⚠️ " + conflict.message);
        warningLabel->setStyleSheet(
            "QLabel { "
            "color: #d32f2f; "
            "background-color: #ffebee; "
            "border: 1px solid #ef5350; "
            "border-radius: 4px; "
            "padding: 8px; "
            "}"
        );
        warningLabel->setWordWrap(true);
        mainLayout->addWidget(warningLabel);
    }

    // 按钮
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    buttonLayout->addStretch();
    QPushButton* okButton = new QPushButton("确定");
    QPushButton* cancelButton = new QPushButton("取消");

    okButton->setDefault(true);
    okButton->setMinimumWidth(80);
    cancelButton->setMinimumWidth(80);

    buttonLayout->addWidget(okButton);
    buttonLayout->addWidget(cancelButton);
    mainLayout->addLayout(buttonLayout);

    connect(okButton, &QPushButton::clicked, &dialog, &QDialog::accept);
    connect(cancelButton, &QPushButton::clicked, &dialog, &QDialog::reject);

    // 显示对话框
    if (dialog.exec() == QDialog::Accepted) {
        count = countSpinBox->value();
        insertCol = beforeRadio->isChecked() ? currentCol : currentCol + 1;
        return true;
    }

    return false;
}


