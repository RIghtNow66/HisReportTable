#ifndef ENHANCEDTABLEVIEW_H
#define ENHANCEDTABLEVIEW_H

#include <QTableView>
#include <QPainter>
#include <QStyleOption>

class ReportDataModel;

class EnhancedTableView : public QTableView
{
    Q_OBJECT

public:
    explicit EnhancedTableView(QWidget* parent = nullptr);

    // 更新合并单元格显示
    void updateSpans();

protected:
    void paintEvent(QPaintEvent* event) override;

private:
    void drawBorders(QPainter* painter);
    void setSpan(int row, int column, int rowSpanCount, int columnSpanCount);
    ReportDataModel* getReportModel() const;
};

#endif // ENHANCEDTABLEVIEW_H