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

class FormulaEngine;
class DayReportParser;

class ReportDataModel : public QAbstractTableModel
{
    Q_OBJECT

public:
    explicit ReportDataModel(QObject* parent = nullptr);
    ~ReportDataModel();

    enum WorkMode {
        REALTIME_MODE,    // 实时单元格绑定模式
        HISTORY_MODE      // 历史报表生成模式
    };

    //  添加模式管理接口
    bool hasHistoryData() const { return !m_fullTimeAxis.isEmpty(); }
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
    bool loadFromExcel(const QString& fileName);
    bool saveToExcel(const QString& fileName);

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

    // 数据绑定
    void resolveDataBindings();

    QFont ensureFontAvailable(const QFont& requestedFont) const;

signals:
    void cellChanged(int row, int col);

private:
    QHash<QPoint, CellData*> m_cells;        // 改为CellData*
    int m_maxRow;
    int m_maxCol;
    FormulaEngine* m_formulaEngine;
    QVector<double> m_rowHeights;
    QVector<double> m_columnWidths;

    GlobalDataConfig m_globalConfig;         // 全局配置

    WorkMode m_currentMode;                                   // 当前工作模式
    QString m_reportName;                                     // 报表名称
    QVector<QDateTime> m_fullTimeAxis;                        // 完整时间轴
    QHash<QString, QVector<double>> m_fullAlignedData;        // 对齐后的数据
};

#endif // REPORTDATAMODEL_H