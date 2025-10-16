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

    // ���ºϲ���Ԫ����ʾ
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
    // ������س�Ա����
    qreal m_zoomFactor;          // ��ǰ���ű���
    int m_baseFontSize;          // ���������С
    int m_baseRowHeight;         // �����и�
    int m_baseColumnWidth;       // �����п�

    // ���������еĳ�ʼ���
    QMap<int, int> m_baseColumnWidths;

	// ���嶯��
    QPropertyAnimation* m_fontAnimation;
    int m_currentAnimatedFontSize;


    // ���ŷ�Χ����
    static constexpr qreal MIN_ZOOM = 0.5;
    static constexpr qreal MAX_ZOOM = 3.0;
    static constexpr qreal ZOOM_STEP = 0.1;
};

#endif // ENHANCEDTABLEVIEW_H