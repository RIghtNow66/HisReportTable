#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QToolBar>
#include <QLineEdit>
#include <QLabel>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QAction>
#include <QMenu>
#include <QDialog>
#include <QPushButton>
#include <QSortFilterProxyModel>
#include <QDateTime> 
#include <QGroupBox>
#include <QRadioButton>
#include <QSpinBox>
#include <QProgressDialog>
            

#include "TimeSettingsDialog.h"
#include "DataBindingConfig.h" 
#include "EnhancedTableView.h"

class ReportDataModel;

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget* parent = nullptr);
    ~MainWindow();

private slots:
    // 文件操作
    void onImportExcel();
    void onExportExcel();

    // 工具操作
    void onFind();
    void onFindNext();
    void onFilter();
    void onClearFilter();

    void onCurrentCellChanged(const QModelIndex& current, const QModelIndex& previous);
    void onFormulaEditFinished();
    void onFormulaTextChanged(); // 监听公式编辑变化
    void onCellClicked(const QModelIndex& index); // 处理单元格点击
    void onCellChanged(int row, int col);

    void onInsertRow();
    void onInsertColumn();
    void onDeleteRow();
    void onDeleteColumn();

    void onRefreshData();  // 保留，但实现会改变
    void onRestoreConfig();

    void onFillDownFormula();

    // 模式切换响应
    void onEditModeChanged(bool editMode);

    void onUnifiedQueryCompleted(bool success, QString message);
    void onUnifiedQueryCanceled();

private:
    // UI组件
    QWidget* m_centralWidget;
    QVBoxLayout* m_mainLayout;

    // 工具栏
    QToolBar* m_toolBar;

    // 公式栏
    QWidget* m_formulaWidget;
    QLabel* m_cellNameLabel;
    QLineEdit* m_formulaEdit;

    // 表格
    EnhancedTableView* m_tableView;
    ReportDataModel* m_dataModel;
    QSortFilterProxyModel* m_filterModel;

    // 查找对话框
    QDialog* m_findDialog;
    QLineEdit* m_findLineEdit;

    // 右键菜单
    QMenu* m_contextMenu;  

    // 状态
    QModelIndex m_currentIndex;
    bool m_updating;
    bool m_formulaEditMode; // 新增：标识是否在公式编辑模式
    QModelIndex m_formulaEditingIndex; // 新增：记录正在编辑公式的单元格

    struct MergeConflictInfo {
        bool hasConflict;
        QString message;
        int safePosition;
    };

    QAction* m_insertRowAction;
    QAction* m_insertColAction;
    QAction* m_deleteRowAction;
    QAction* m_deleteColAction;
    QAction* m_fillFormulaAction;

    TimeSettingsDialog* m_timeSettingsDialog;

    QProgressDialog* m_unifiedQueryProgress = nullptr;

private:
    void setupUI();
    void setupToolBar();
    void setupFormulaBar();
    void setupTableView();
    void setupContextMenu();
    void setupFindDialog();
    void updateFormulaBar(const QModelIndex& index);

    void enterFormulaEditMode();
    void exitFormulaEditMode();
    bool isInFormulaEditMode() const;

    void applyRowColumnSizes();

    QString adjustFormulaReferences(const QString& formula, int rowOffset);

    int findFillEndRow(int currentRow, int currentCol);

    void exportData();
    void exportTemplate();
    QString generateFileName(const QString& suffix);

    // 冲突检测方法
    MergeConflictInfo checkRowInsertConflict(int row);
    MergeConflictInfo checkColumnInsertConflict(int col);

    // 插入对话框方法
    bool showInsertRowDialog(int currentRow, int& insertRow, int& count);
    bool showInsertColumnDialog(int currentCol, int& insertCol, int& count);

    // 模式重新设置函数
    void updateUIForEditMode(bool editMode);

};

#endif // MAINWINDOW_H
