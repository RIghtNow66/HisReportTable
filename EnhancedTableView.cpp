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
                qDebug() << " 已达到最小缩放限制 (50%)";
                emit zoomChanged(m_zoomFactor);  // 触发实时反馈
            }
            else if (newZoom > MAX_ZOOM) {
                newZoom = MAX_ZOOM;
                qDebug() << " 已达到最大缩放限制 (300%)";
                emit zoomChanged(m_zoomFactor);  // 触发实时反馈
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

    // 使用平滑动画
    applySmoothZoom();

    // 发射信号用于实时反馈
    emit zoomChanged(m_zoomFactor);
}


// 重置缩放到100%
void EnhancedTableView::resetZoom()
{
    setZoomFactor(1.0);
    qDebug() << "缩放已重置为100%";
}

void EnhancedTableView::keyPressEvent(QKeyEvent* event)
{
    if (event->modifiers() & Qt::ControlModifier) {
        if (event->key() == Qt::Key_Plus || event->key() == Qt::Key_Equal) {
            qreal newZoom = m_zoomFactor + ZOOM_STEP;
            if (newZoom > MAX_ZOOM) {
                qDebug() << "已达到最大缩放限制 (300%)";
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
                qDebug() << "已达到最小缩放限制 (50%)";
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

    // 如果正在动画中，停止当前动画
    if (m_fontAnimation && m_fontAnimation->state() == QAbstractAnimation::Running) {
        m_fontAnimation->stop();
    }

    // 创建字体大小动画
    if (!m_fontAnimation) {
        m_fontAnimation = new QPropertyAnimation(this, "animatedFontSize", this);
        m_fontAnimation->setDuration(150);  // 150ms 平滑过渡
        m_fontAnimation->setEasingCurve(QEasingCurve::OutCubic);  // 平滑曲线
    }

    m_fontAnimation->setStartValue(m_currentAnimatedFontSize);
    m_fontAnimation->setEndValue(targetFontSize);
    m_fontAnimation->start();
}

void EnhancedTableView::setAnimatedFontSize(int size)
{
    m_currentAnimatedFontSize = size;

    // 应用字体大小
    QFont newFont = font();
    newFont.setPointSize(size);
    setFont(newFont);

    // 同步调整行高（按相同比例）
    qreal currentZoom = static_cast<qreal>(size) / m_baseFontSize;
    int newRowHeight = qRound(m_baseRowHeight * currentZoom);
    verticalHeader()->setDefaultSectionSize(qMax(15, newRowHeight));

    // 同步调整列宽
    adjustAllColumnsWidth();

    // 调整表头字体
    QFont headerFont = horizontalHeader()->font();
    headerFont.setPointSize(size);
    horizontalHeader()->setFont(headerFont);
    verticalHeader()->setFont(headerFont);

    // 刷新显示
    viewport()->update();
}

void EnhancedTableView::adjustAllColumnsWidth()
{
    if (!model()) {
        return;
    }

    int columnCount = model()->columnCount();

    // 第一次调用时，保存所有列的初始宽度
    if (m_baseColumnWidths.isEmpty()) {
        for (int col = 0; col < columnCount; ++col) {
            int width = columnWidth(col);
            if (width > 0) {
                m_baseColumnWidths[col] = width;
            }
            else {
                m_baseColumnWidths[col] = m_baseColumnWidth;  // 使用默认值
            }
        }
    }

    //  按缩放比例调整所有列宽
    for (int col = 0; col < columnCount; ++col) {
        if (m_baseColumnWidths.contains(col)) {
            int baseWidth = m_baseColumnWidths[col];
            int newWidth = qRound(baseWidth * m_zoomFactor);
            setColumnWidth(col, qMax(30, newWidth));  // 最小30像素
        }
    }
}