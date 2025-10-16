#ifndef ENHANCEDTABLEVIEW_H
#define ENHANCEDTABLEVIEW_H

#include <QTableView>
#include <QPainter>
#include <QStyleOption>
#include <QWheelEvent>
#include <QKeyEvent> 
#include <QPropertyAnimation>
#include <QStyledItemDelegate>

class ReportDataModel;

class ScaledFontDelegate : public QStyledItemDelegate
{
    Q_OBJECT
public:
    explicit ScaledFontDelegate(QObject* parent = nullptr)
        : QStyledItemDelegate(parent)
        , m_useScaledFont(false)
        , m_scaledFontSize(10)
    {
    }

    void setScaledFontSize(int size) {
        m_scaledFontSize = size;
        m_useScaledFont = (size > 0);
    }

    void resetScaling() {
        m_useScaledFont = false;
    }

protected:
    void initStyleOption(QStyleOptionViewItem* option, const QModelIndex& index) const override
    {
        QStyledItemDelegate::initStyleOption(option, index);

        if (m_useScaledFont) {
            QFont font = option->font;
            font.setPointSize(m_scaledFontSize);
            option->font = font;
        }
    }

private:
    bool m_useScaledFont;
    int m_scaledFontSize;
};

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

    void saveBaseRowHeights();
    void resetColumnWidthsBase();

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

    QMap<int, int> m_baseRowHeights;
    QMap<int, int> m_baseColumnWidths;

	// 字体动画
    QPropertyAnimation* m_fontAnimation;
    int m_currentAnimatedFontSize;


    // 缩放范围常量
    static constexpr qreal MIN_ZOOM = 0.5;
    static constexpr qreal MAX_ZOOM = 3.0;
    static constexpr qreal ZOOM_STEP = 0.1;

    ScaledFontDelegate* m_fontDelegate;
};



#endif // ENHANCEDTABLEVIEW_H