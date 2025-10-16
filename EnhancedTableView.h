#ifndef ENHANCEDTABLEVIEW_H
#define ENHANCEDTABLEVIEW_H

#include <QTableView>
#include <QPainter>
#include <QStyleOption>
#include <QWheelEvent>
#include <QKeyEvent> 
#include <QPropertyAnimation>

class ReportDataModel;

class EnhancedTableView : public QTableView
{
    Q_OBJECT
    Q_PROPERTY(int animatedFontSize READ animatedFontSize WRITE setAnimatedFontSize)

public:
    explicit EnhancedTableView(QWidget* parent = nullptr);

    // 更新合并单元格显示
    void updateSpans();

    void setZoomFactor(qreal factor);
    qreal getZoomFactor() const { return m_zoomFactor; }
    void resetZoom();

    int animatedFontSize() const { return m_currentAnimatedFontSize; }
    void setAnimatedFontSize(int size);

signals:
    void zoomChanged(qreal factor);

protected:
    void paintEvent(QPaintEvent* event) override;
    void wheelEvent(QWheelEvent* event) override;

    void keyPressEvent(QKeyEvent* event) override;

private:
    void drawBorders(QPainter* painter);
    void setSpan(int row, int column, int rowSpanCount, int columnSpanCount);
    ReportDataModel* getReportModel() const;

    void applySmoothZoom();

    void adjustAllColumnsWidth();

private:
    // 缩放相关成员变量
    qreal m_zoomFactor;          // 当前缩放比例
    int m_baseFontSize;          // 基础字体大小
    int m_baseRowHeight;         // 基础行高
    int m_baseColumnWidth;       // 基础列宽

    // 保存所有列的初始宽度
    QMap<int, int> m_baseColumnWidths;

	// 字体动画
    QPropertyAnimation* m_fontAnimation;
    int m_currentAnimatedFontSize;


    // 缩放范围常量
    static constexpr qreal MIN_ZOOM = 0.5;
    static constexpr qreal MAX_ZOOM = 3.0;
    static constexpr qreal ZOOM_STEP = 0.1;
};

#endif // ENHANCEDTABLEVIEW_H