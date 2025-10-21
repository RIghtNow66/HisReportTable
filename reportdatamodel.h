#ifndef REPORTDATAMODEL_H
#define REPORTDATAMODEL_H

#include "DataBindingConfig.h"
#include <QHash>
#include <QAbstractTableModel>
#include <QFontInfo>      // 添加这个
#include <QFontDatabase>  // 如果需要的话
#include <QBrush>         // 添加这个
#include <QPoint>
#include <QSize>
#include <QVector> 
#include <QProgressDialog>

// qHash 函数必须在 QHash 使用之前定义
inline uint qHash(const QPoint& key, uint seed = 0) noexcept
{
    return qHash(qMakePair(key.x(), key.y()), seed);  // 改用 qMakePair
}

class BaseReportParser;
class MonthReportParser;
class DayReportParser;
class FormulaEngine;
class QProgressDialog;

class ReportDataModel : public QAbstractTableModel
{
    Q_OBJECT

public:
    enum ChangeType {
        NO_CHANGE,           // 无变化
        FORMULA_ONLY,        // 只有公式变化
        BINDING_ONLY,        // 只有绑定变化
        MIXED_CHANGE         // 混合变化
    };

    enum ExportMode {
        EXPORT_DATA,
        EXPORT_TEMPLATE
    };

    ChangeType detectChanges();  // 检测变化类型
    void saveRefreshSnapshot();  // 保存刷新快照
    bool isFirstRefresh() const { return m_isFirstRefresh; }

    enum ReportType {
        NORMAL_EXCEL, // 普通Excel或未识别
        DAY_REPORT,   // 日报
        MONTH_REPORT  // 为月报预留
    };


    explicit ReportDataModel(QObject* parent = nullptr);
    ~ReportDataModel();

    bool loadReportTemplate(const QString& fileName); // 替代 loadFromExcel 和 loadDayReport
    bool refreshReportData(QProgressDialog* progress); // 修改签名
    void restoreToTemplate(); // 新增函数

    ReportType getReportType() const;
    BaseReportParser* getParser() const;
    void notifyDataChanged();

    //  添加模式管理接口
    QString getReportName() const { return m_reportName; }
    bool hasDataBindings() const;  //  检查是否有##绑定


    // Qt Model 接口
    int rowCount(const QModelIndex& parent = QModelIndex()) const override;
    int columnCount(const QModelIndex& parent = QModelIndex()) const override;
    QVariant data(const QModelIndex& index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex& index, const QVariant& value, int role = Qt::EditRole) override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    QSize span(const QModelIndex& index) const;

    // 行列操作
    bool insertRows(int row, int count, const QModelIndex& parent = QModelIndex()) override;
    bool removeRows(int row, int count, const QModelIndex& parent = QModelIndex()) override;
    bool insertColumns(int column, int count, const QModelIndex& parent = QModelIndex()) override;
    bool removeColumns(int column, int count, const QModelIndex& parent = QModelIndex()) override;

    // 文件操作
    bool saveToExcel(const QString& fileName, ExportMode mode = EXPORT_DATA);

    // 单元格访问
    void clearAllCells();
    void addCellDirect(int row, int col, CellData* cell);  // 改为CellData*
    void updateModelSize(int newRowCount, int newColCount);
    const QHash<QPoint, CellData*>& getAllCells() const;   // 改为CellData*
    void recalculateAllFormulas();
    const CellData* getCell(int row, int col) const;       // 改为CellData*
    CellData* getCell(int row, int col);                   // 改为CellData*
    CellData* ensureCell(int row, int col);                // 改为CellData*
    void calculateFormula(int row, int col);
    QString cellAddress(int row, int col) const;

    // 行高列宽
    void setRowHeight(int row, double height);
    double getRowHeight(int row) const;
    void setColumnWidth(int col, double width);
    double getColumnWidth(int col) const;
    const QVector<double>& getAllRowHeights() const;
    const QVector<double>& getAllColumnWidths() const;
    void clearSizes();

    QFont ensureFontAvailable(const QFont& requestedFont) const;

    bool hasExecutedQueries() const;

    // 模式管理接口
    void setEditMode(bool editMode);
    bool isEditMode() const { return m_editMode; }


signals:
    void cellChanged(int row, int col);

    // 模式变化信号 
    void editModeChanged(bool editMode);

private:
    bool loadFromExcelFile(const QString& fileName);

    QSet<QString> getCurrentBindings() const;     // 获取当前所有绑定标记
    QSet<QPoint> getCurrentFormulas() const;      // 获取当前所有公式位置
    QList<QString> getNewBindings() const;        // 获取新增的绑定标记

    // ===== 公式依赖检查 =====
    bool checkFormulaDependenciesReady(int row, int col);
    QPoint parseCellReference(const QString& cellRef) const;

private:
    QHash<QPoint, CellData*> m_cells;        // 改为CellData*
    int m_maxRow;
    int m_maxCol;
    FormulaEngine* m_formulaEngine;
    QVector<double> m_rowHeights;
    QVector<double> m_columnWidths;

    GlobalDataConfig m_globalConfig;         // 全局配置

    QString m_reportName;                                     // 报表名称

    ReportType m_reportType;
    BaseReportParser* m_parser;

    struct RefreshSnapshot {
        QSet<QString> bindingKeys;     // 上次刷新时的绑定标记集合
        QSet<QPoint> formulaCells;     // 上次刷新时的公式单元格位置
        QSet<QPoint> dataMarkerCells;
        bool isEmpty() const { return bindingKeys.isEmpty() && formulaCells.isEmpty(); dataMarkerCells.isEmpty();
        }
    };

    RefreshSnapshot m_lastSnapshot;    // 上次刷新的快照
    bool m_isFirstRefresh = true;      // 是否首次刷新

    bool m_editMode = true;  // 默认为编辑模式

};

#endif // REPORTDATAMODEL_H