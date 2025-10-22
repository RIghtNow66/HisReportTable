#ifndef REPORTDATAMODEL_H
#define REPORTDATAMODEL_H

#include "DataBindingConfig.h"
#include <QHash>
#include <QAbstractTableModel>
#include <QFontInfo>
#include <QFontDatabase>
#include <QBrush>
#include <QPoint>
#include <QSize>
#include <QVector> 
#include <QProgressDialog>

inline uint qHash(const QPoint& key, uint seed = 0) noexcept
{
    return qHash(qMakePair(key.x(), key.y()), seed);
}

class BaseReportParser;
class MonthReportParser;
class DayReportParser;
class UnifiedQueryParser;  // 新增前向声明
class FormulaEngine;
class QProgressDialog;

struct HistoryReportConfig;
struct TimeRangeConfig;

class ReportDataModel : public QAbstractTableModel
{
    Q_OBJECT

public:
    // ===== 报表模式枚举 =====
    enum ReportMode {
        TEMPLATE_MODE,
        UNIFIED_QUERY_MODE
    };

    enum TemplateType {
        NORMAL_EXCEL,
        DAY_REPORT,
        MONTH_REPORT
    };

    enum ChangeType {
        NO_CHANGE,
        FORMULA_ONLY,
        BINDING_ONLY,
        MIXED_CHANGE
    };

    enum ExportMode {
        EXPORT_DATA,
        EXPORT_TEMPLATE
    };

    enum UnifiedQueryChangeType {
        UQ_NO_CHANGE,
        UQ_FORMULA_ONLY,
        UQ_NEED_REQUERY
    };

public:
    // ===== 模式管理接口 =====
    ReportMode currentMode() const { return m_currentMode; }
    bool isUnifiedQueryMode() const { return m_currentMode == UNIFIED_QUERY_MODE; }

    // ===== 统一查询模式接口 =====
    bool hasUnifiedQueryData() const;  // 新增：判断是否已查询数据
    void setTimeRangeForQuery(const TimeRangeConfig& config);  // 新增：设置时间范围
    bool exportConfigFile(const QString& fileName);  // 保留：导出配置

    ChangeType detectChanges();
    void saveRefreshSnapshot();
    bool isFirstRefresh() const { return m_isFirstRefresh; }

    explicit ReportDataModel(QObject* parent = nullptr);
    ~ReportDataModel();

    bool loadReportTemplate(const QString& fileName);
    bool refreshReportData(QProgressDialog* progress);
    void restoreToTemplate();

    TemplateType getReportType() const;
    BaseReportParser* getParser() const;
    void notifyDataChanged();

    QString getReportName() const { return m_reportName; }
    bool hasDataBindings() const;

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
    void addCellDirect(int row, int col, CellData* cell);
    void updateModelSize(int newRowCount, int newColCount);
    const QHash<QPoint, CellData*>& getAllCells() const;
    void recalculateAllFormulas();
    const CellData* getCell(int row, int col) const;
    CellData* getCell(int row, int col);
    CellData* ensureCell(int row, int col);
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

    void optimizeMemory();
    void markFormulaDirty(int row, int col);
    void markDependentFormulasDirty(int changedRow, int changedCol);

    void updateEditability();  // 更新可编辑状态
    int getDataColumnCount() const { return m_dataColumnCount; }  // 获取数据列数

    UnifiedQueryChangeType detectUnifiedQueryChanges();

signals:
    void cellChanged(int row, int col);
    void editModeChanged(bool editMode);

private:
    QHash<QPoint, CellData*> m_cells;
    int m_maxRow;
    int m_maxCol;
    FormulaEngine* m_formulaEngine;
    QVector<double> m_rowHeights;
    QVector<double> m_columnWidths;

    GlobalDataConfig m_globalConfig;
    QString m_reportName;

    TemplateType m_reportType;
    BaseReportParser* m_parser;

    struct RefreshSnapshot {
        QSet<QString> bindingKeys;
        QSet<QPoint> formulaCells;
        QSet<QPoint> dataMarkerCells;
        bool isEmpty() const {
            return bindingKeys.isEmpty() && formulaCells.isEmpty() && dataMarkerCells.isEmpty();
        }
    };

    RefreshSnapshot m_lastSnapshot;
    bool m_isFirstRefresh = true;
    bool m_editMode = true;
    QSet<QPoint> m_dirtyFormulas;

    // ===== 模式管理变量 =====
    ReportMode m_currentMode;
    TemplateType m_templateType;

    int m_dataColumnCount = 0;  // 新增：记录统一查询模式下数据列数

private:
    // ===== 模式分发函数 =====
    QVariant getTemplateCellData(const QModelIndex& index, int role) const;
    QVariant getUnifiedQueryCellData(const QModelIndex& index, int role) const;  // 修改：实现

    bool refreshTemplateReport(QProgressDialog* progress);

    // ===== 统一查询辅助函数 =====
    bool loadUnifiedQueryConfig(const QString& filePath);  // 修改：实现
    bool refreshUnifiedQuery(QProgressDialog* progress);   // 修改：实现

    bool loadFromExcelFile(const QString& fileName);

    QSet<QString> getCurrentBindings() const;
    QSet<QPoint> getCurrentFormulas() const;
    QList<QString> getNewBindings() const;

    // ===== 公式依赖检查 =====
    bool checkFormulaDependenciesReady(int row, int col);
    QPoint parseCellReference(const QString& cellRef) const;
    bool detectCircularDependency(int row, int col, QSet<QPoint>& visitedCells);

    Qt::ItemFlags getTemplateModeFlags(const QModelIndex& index) const;
    Qt::ItemFlags getUnifiedQueryModeFlags(const QModelIndex& index) const;

    void restoreTemplateReport();  // 拆分：模板模式还原
    void restoreUnifiedQuery();    // 拆分：统一查询还原

};

#endif // REPORTDATAMODEL_H