#include "EnhancedTableView.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"
#include <QPaintEvent>
#include <QPainter>
#include <QHeaderView>
#include <QAbstractProxyModel>
#include <QPen>
#include <QSet> 
#include <qDebug>

EnhancedTableView::EnhancedTableView(QWidget* parent)
    : QTableView(parent)
    , m_zoomFactor(1.0)          
    , m_baseFontSize(9)          
    , m_baseRowHeight(25)        
    , m_baseColumnWidth(80)      
    , m_fontAnimation(nullptr)         
    , m_currentAnimatedFontSize(9)
{
    QFont currentFont = font();
    if (currentFont.pointSize() > 0) {
        m_baseFontSize = currentFont.pointSize();
        m_currentAnimatedFontSize = m_baseFontSize;
    }

    if (verticalHeader()->defaultSectionSize() > 0) {
        m_baseRowHeight = verticalHeader()->defaultSectionSize();
    }
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

            // ========== �������޸�����ʼ ==========
            // ���ںϲ���Ԫ�񣬷ֱ��ȡ�ĸ���Ե�ı߿���Ϣ
            RTCellBorder topBorder, bottomBorder, leftBorder, rightBorder;

            if (masterCell->mergedRange.isMerged()) {
                // �ϲ���Ԫ�񣺴��ĸ���Ե��Ԫ��ֱ��ȡ�߿�
                int startRow = masterCell->mergedRange.startRow;
                int endRow = masterCell->mergedRange.endRow;
                int startCol = masterCell->mergedRange.startCol;
                int endCol = masterCell->mergedRange.endCol;

                // �ϱ߿�ȡ��һ�е��ϱ߿�
                const CellData* topCell = reportModel->getCell(startRow, startCol);
                topBorder = topCell ? topCell->style.border : RTCellBorder();

                // �±߿�ȡ���һ�е��±߿�
                const CellData* bottomCell = reportModel->getCell(endRow, startCol);
                bottomBorder = bottomCell ? bottomCell->style.border : RTCellBorder();

                // ��߿�ȡ��һ�е���߿�
                const CellData* leftCell = reportModel->getCell(startRow, startCol);
                leftBorder = leftCell ? leftCell->style.border : RTCellBorder();

                // �ұ߿�ȡ���һ�е��ұ߿򣨹ؼ�����
                const CellData* rightCell = reportModel->getCell(startRow, endCol);
                rightBorder = rightCell ? rightCell->style.border : RTCellBorder();
            }
            else {
                // ��ͨ��Ԫ��ֱ��ʹ���Լ��ı߿�
                topBorder = bottomBorder = leftBorder = rightBorder = masterCell->style.border;
            }

            // �����ĸ��߿�ʹ�÷ֱ��ȡ�ı߿���Ϣ��
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
            // ========== �������޸������� ==========

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

void EnhancedTableView::wheelEvent(QWheelEvent* event)
{
    if (event->modifiers() & Qt::ControlModifier) {
        int delta = event->angleDelta().y();

        if (delta != 0) {
            qreal zoomDelta = (delta > 0) ? ZOOM_STEP : -ZOOM_STEP;
            qreal newZoom = m_zoomFactor + zoomDelta;
            newZoom = qBound(MIN_ZOOM, newZoom, MAX_ZOOM);

            if (newZoom < MIN_ZOOM) {
                newZoom = MIN_ZOOM;
                qDebug() << " �Ѵﵽ��С�������� (50%)";
                emit zoomChanged(m_zoomFactor);  // ����ʵʱ����
            }
            else if (newZoom > MAX_ZOOM) {
                newZoom = MAX_ZOOM;
                qDebug() << " �Ѵﵽ����������� (300%)";
                emit zoomChanged(m_zoomFactor);  // ����ʵʱ����
            }
            else if (newZoom != m_zoomFactor) {
                setZoomFactor(newZoom);
            }
        }

        event->accept();
    }
    else {
        QTableView::wheelEvent(event);
    }
}

void EnhancedTableView::setZoomFactor(qreal factor)
{
    factor = qBound(MIN_ZOOM, factor, MAX_ZOOM);

    if (qAbs(m_zoomFactor - factor) < 0.01) {
        return;
    }

    m_zoomFactor = factor;

    // ʹ��ƽ������
    applySmoothZoom();

    // �����ź�����ʵʱ����
    emit zoomChanged(m_zoomFactor);
}


// �������ŵ�100%
void EnhancedTableView::resetZoom()
{
    setZoomFactor(1.0);
    qDebug() << "����������Ϊ100%";
}

void EnhancedTableView::keyPressEvent(QKeyEvent* event)
{
    if (event->modifiers() & Qt::ControlModifier) {
        if (event->key() == Qt::Key_Plus || event->key() == Qt::Key_Equal) {
            qreal newZoom = m_zoomFactor + ZOOM_STEP;
            if (newZoom > MAX_ZOOM) {
                qDebug() << "�Ѵﵽ����������� (300%)";
                emit zoomChanged(m_zoomFactor);
            }
            else {
                setZoomFactor(newZoom);
            }
            event->accept();
            return;
        }
        else if (event->key() == Qt::Key_Minus) {
            qreal newZoom = m_zoomFactor - ZOOM_STEP;
            if (newZoom < MIN_ZOOM) {
                qDebug() << "�Ѵﵽ��С�������� (50%)";
                emit zoomChanged(m_zoomFactor);
            }
            else {
                setZoomFactor(newZoom);
            }
            event->accept();
            return;
        }
        else if (event->key() == Qt::Key_0) {
            resetZoom();
            event->accept();
            return;
        }
    }

    QTableView::keyPressEvent(event);
}

void EnhancedTableView::applySmoothZoom()
{
    int targetFontSize = qRound(m_baseFontSize * m_zoomFactor);
    targetFontSize = qMax(6, targetFontSize);

    // ������ڶ����У�ֹͣ��ǰ����
    if (m_fontAnimation && m_fontAnimation->state() == QAbstractAnimation::Running) {
        m_fontAnimation->stop();
    }

    // ���������С����
    if (!m_fontAnimation) {
        m_fontAnimation = new QPropertyAnimation(this, "animatedFontSize", this);
        m_fontAnimation->setDuration(150);  // 150ms ƽ������
        m_fontAnimation->setEasingCurve(QEasingCurve::OutCubic);  // ƽ������
    }

    m_fontAnimation->setStartValue(m_currentAnimatedFontSize);
    m_fontAnimation->setEndValue(targetFontSize);
    m_fontAnimation->start();
}

void EnhancedTableView::setAnimatedFontSize(int size)
{
    m_currentAnimatedFontSize = size;

    // Ӧ�������С
    QFont newFont = font();
    newFont.setPointSize(size);
    setFont(newFont);

    // ͬ�������иߣ�����ͬ������
    qreal currentZoom = static_cast<qreal>(size) / m_baseFontSize;
    int newRowHeight = qRound(m_baseRowHeight * currentZoom);
    verticalHeader()->setDefaultSectionSize(qMax(15, newRowHeight));

    // ͬ�������п�
    adjustAllColumnsWidth();

    // ������ͷ����
    QFont headerFont = horizontalHeader()->font();
    headerFont.setPointSize(size);
    horizontalHeader()->setFont(headerFont);
    verticalHeader()->setFont(headerFont);

    // ˢ����ʾ
    viewport()->update();
}

void EnhancedTableView::adjustAllColumnsWidth()
{
    if (!model()) {
        return;
    }

    int columnCount = model()->columnCount();

    // ��һ�ε���ʱ�����������еĳ�ʼ���
    if (m_baseColumnWidths.isEmpty()) {
        for (int col = 0; col < columnCount; ++col) {
            int width = columnWidth(col);
            if (width > 0) {
                m_baseColumnWidths[col] = width;
            }
            else {
                m_baseColumnWidths[col] = m_baseColumnWidth;  // ʹ��Ĭ��ֵ
            }
        }
    }

    //  �����ű������������п�
    for (int col = 0; col < columnCount; ++col) {
        if (m_baseColumnWidths.contains(col)) {
            int baseWidth = m_baseColumnWidths[col];
            int newWidth = qRound(baseWidth * m_zoomFactor);
            setColumnWidth(col, qMax(30, newWidth));  // ��С30����
        }
    }
}