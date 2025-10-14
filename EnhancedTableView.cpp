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

            QPoint masterCellPos(row, col);

            if (cell->mergedRange.isMerged()) {
                masterCellPos.setX(cell->mergedRange.startRow);
                masterCellPos.setY(cell->mergedRange.startCol);
            }

            if (drawnSpans.contains(masterCellPos)) {
                continue;
            }

            QModelIndex masterIndex = reportModel->index(masterCellPos.x(), masterCellPos.y());
            QRect cellRect = visualRect(masterIndex);

            const CellData* masterCell = reportModel->getCell(masterCellPos.x(), masterCellPos.y());
            if (!masterCell) continue;

            // ========== 【核心修复】开始 ==========
            // 对于合并单元格，分别获取四个边缘的边框信息
            RTCellBorder topBorder, bottomBorder, leftBorder, rightBorder;

            if (masterCell->mergedRange.isMerged()) {
                // 合并单元格：从四个边缘单元格分别读取边框
                int startRow = masterCell->mergedRange.startRow;
                int endRow = masterCell->mergedRange.endRow;
                int startCol = masterCell->mergedRange.startCol;
                int endCol = masterCell->mergedRange.endCol;

                // 上边框：取第一行的上边框
                const CellData* topCell = reportModel->getCell(startRow, startCol);
                topBorder = topCell ? topCell->style.border : RTCellBorder();

                // 下边框：取最后一行的下边框
                const CellData* bottomCell = reportModel->getCell(endRow, startCol);
                bottomBorder = bottomCell ? bottomCell->style.border : RTCellBorder();

                // 左边框：取第一列的左边框
                const CellData* leftCell = reportModel->getCell(startRow, startCol);
                leftBorder = leftCell ? leftCell->style.border : RTCellBorder();

                // 右边框：取最后一列的右边框（关键！）
                const CellData* rightCell = reportModel->getCell(startRow, endCol);
                rightBorder = rightCell ? rightCell->style.border : RTCellBorder();
            }
            else {
                // 普通单元格：直接使用自己的边框
                topBorder = bottomBorder = leftBorder = rightBorder = masterCell->style.border;
            }

            // 绘制四个边框（使用分别获取的边框信息）
            if (topBorder.top != RTBorderStyle::None) {
                QPen pen(topBorder.topColor, static_cast<int>(topBorder.top));
                painter->setPen(pen);
                painter->drawLine(cellRect.topLeft(), cellRect.topRight());
            }
            if (bottomBorder.bottom != RTBorderStyle::None) {
                QPen pen(bottomBorder.bottomColor, static_cast<int>(bottomBorder.bottom));
                painter->setPen(pen);
                painter->drawLine(cellRect.bottomLeft(), cellRect.bottomRight());
            }
            if (leftBorder.left != RTBorderStyle::None) {
                QPen pen(leftBorder.leftColor, static_cast<int>(leftBorder.left));
                painter->setPen(pen);
                painter->drawLine(cellRect.topLeft(), cellRect.bottomLeft());
            }
            if (rightBorder.right != RTBorderStyle::None) {
                QPen pen(rightBorder.rightColor, static_cast<int>(rightBorder.right));
                painter->setPen(pen);
                painter->drawLine(cellRect.topRight(), cellRect.bottomRight());
            }
            // ========== 【核心修复】结束 ==========

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