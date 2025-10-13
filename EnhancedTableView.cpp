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
    // �ȵ��û���Ļ���
    QTableView::paintEvent(event);

    // Ȼ������Զ���߿�
    QPainter painter(viewport());
    drawBorders(&painter);
}

void EnhancedTableView::drawBorders(QPainter* painter)
{
    ReportDataModel* reportModel = getReportModel();
    if (!reportModel) return;

    painter->save();

    // ���ڸ����Ѿ����ƹ��ĺϲ���Ԫ������Ͻ�����
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

            // Ĭ������£���Ԫ��ġ�����Ԫ�񡱾������Լ�
            QPoint masterCellPos(row, col);

            // �����ǰ��Ԫ���Ǻϲ���Ԫ���һ���֣��ҵ��������Ͻ�����Ԫ��
            if (cell->mergedRange.isMerged()) {
                masterCellPos.setX(cell->mergedRange.startRow);
                masterCellPos.setY(cell->mergedRange.startCol);
            }

            // ����������Ԫ�񣨻��������ĺϲ������Ѿ����ƹ��ˣ�������
            if (drawnSpans.contains(masterCellPos)) {
                continue;
            }

            // ��ȡ����Ԫ�������������ռ��������������
            QModelIndex masterIndex = reportModel->index(masterCellPos.x(), masterCellPos.y());
            QRect cellRect = visualRect(masterIndex);

            // ��ȡ����Ԫ�����ʽ�������߿�
            const CellData* masterCell = reportModel->getCell(masterCellPos.x(), masterCellPos.y());
            if (!masterCell) continue;

            const RTCellBorder& border = masterCell->style.border;

            // �������������ĸ���߿�
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

            // �������һ���ϲ����򣬽���������Ԫ����롰�ѻ��ơ�����
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
    // ������ܴ��ڵĴ���ģ��
    QAbstractItemModel* currentModel = model();

    // ����Ǵ���ģ�ͣ���ȡԴģ��
    while (QAbstractProxyModel* proxyModel = qobject_cast<QAbstractProxyModel*>(currentModel)) {
        currentModel = proxyModel->sourceModel();
    }

    return qobject_cast<ReportDataModel*>(currentModel);
}