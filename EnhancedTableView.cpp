#include "EnhancedTableView.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include <QPaintEvent>
#include <QPainter>
#include <QHeaderView>
#include <QAbstractProxyModel>
#include <QPen>
#include <QSet> 

EnhancedTableView::EnhancedTableView(QWidget* parent)
    : QTableView(parent)
{
}

void EnhancedTableView::paintEvent(QPaintEvent* event)
{
    // 先调用基类的绘制
    QTableView::paintEvent(event);

    // 然后绘制自定义边框
    QPainter painter(viewport());
    drawBorders(&painter);
}

void EnhancedTableView::drawBorders(QPainter* painter)
{
    ReportDataModel* reportModel = getReportModel();
    if (!reportModel) return;

    painter->save();

    // 用于跟踪已经绘制过的合并单元格的左上角坐标
    QSet<QPoint> drawnSpans;

    QRect viewportRect = viewport()->rect();
    int firstRow = rowAt(viewportRect.top());
    int lastRow = rowAt(viewportRect.bottom());
    int firstCol = columnAt(viewportRect.left());
    int lastCol = columnAt(viewportRect.right());

    if (firstRow < 0) firstRow = 0;
    if (lastRow < 0) lastRow = reportModel->rowCount() - 1;
    if (firstCol < 0) firstCol = 0;
    if (lastCol < 0) lastCol = reportModel->columnCount() - 1;

    for (int row = firstRow; row <= lastRow; ++row) {
        for (int col = firstCol; col <= lastCol; ++col) {
            const CellData* cell = reportModel->getCell(row, col);
            if (!cell) continue;

            // 默认情况下，单元格的“主单元格”就是它自己
            QPoint masterCellPos(row, col);

            // 如果当前单元格是合并单元格的一部分，找到它的左上角主单元格
            if (cell->mergedRange.isMerged()) {
                masterCellPos.setX(cell->mergedRange.startRow);
                masterCellPos.setY(cell->mergedRange.startCol);
            }

            // 如果这个主单元格（或它所属的合并区域）已经绘制过了，就跳过
            if (drawnSpans.contains(masterCellPos)) {
                continue;
            }

            // 获取主单元格的索引和它所占的完整矩形区域
            QModelIndex masterIndex = reportModel->index(masterCellPos.x(), masterCellPos.y());
            QRect cellRect = visualRect(masterIndex);

            // 获取主单元格的样式来决定边框
            const CellData* masterCell = reportModel->getCell(masterCellPos.x(), masterCellPos.y());
            if (!masterCell) continue;

            const RTCellBorder& border = masterCell->style.border;

            // 绘制这个区域的四个外边框
            if (border.top != RTBorderStyle::None) {
                QPen pen(border.topColor, static_cast<int>(border.top));
                painter->setPen(pen);
                painter->drawLine(cellRect.topLeft(), cellRect.topRight());
            }
            if (border.bottom != RTBorderStyle::None) {
                QPen pen(border.bottomColor, static_cast<int>(border.bottom));
                painter->setPen(pen);
                painter->drawLine(cellRect.bottomLeft(), cellRect.bottomRight());
            }
            if (border.left != RTBorderStyle::None) {
                QPen pen(border.leftColor, static_cast<int>(border.left));
                painter->setPen(pen);
                painter->drawLine(cellRect.topLeft(), cellRect.bottomLeft());
            }
            if (border.right != RTBorderStyle::None) {
                QPen pen(border.rightColor, static_cast<int>(border.right));
                painter->setPen(pen);
                painter->drawLine(cellRect.topRight(), cellRect.bottomRight());
            }

            // 如果这是一个合并区域，将它的主单元格加入“已绘制”集合
            if (masterCell->mergedRange.isMerged()) {
                drawnSpans.insert(masterCellPos);
            }
        }
    }

    painter->restore();
}


void EnhancedTableView::setSpan(int row, int column, int rowSpanCount, int columnSpanCount)
{
    QTableView::setSpan(row, column, rowSpanCount, columnSpanCount);
}

void EnhancedTableView::updateSpans()
{
    ReportDataModel* reportModel = getReportModel();
    if (!reportModel) return;

    clearSpans();

    const auto& allCells = reportModel->getAllCells();
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const QPoint& pos = it.key();
        const CellData* cell = it.value();

        if (cell && cell->mergedRange.isMerged() &&
            pos.x() == cell->mergedRange.startRow &&
            pos.y() == cell->mergedRange.startCol)
        {
            int rowSpan = cell->mergedRange.rowSpan();
            int colSpan = cell->mergedRange.colSpan();
            if (rowSpan > 1 || colSpan > 1) {
                setSpan(pos.x(), pos.y(), rowSpan, colSpan);
            }
        }
    }
}


ReportDataModel* EnhancedTableView::getReportModel() const
{
    // 处理可能存在的代理模型
    QAbstractItemModel* currentModel = model();

    // 如果是代理模型，获取源模型
    while (QAbstractProxyModel* proxyModel = qobject_cast<QAbstractProxyModel*>(currentModel)) {
        currentModel = proxyModel->sourceModel();
    }

    return qobject_cast<ReportDataModel*>(currentModel);
}