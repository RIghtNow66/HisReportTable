#ifndef FORMULAENGINE_H
#define FORMULAENGINE_H

#include <QObject>
#include <QVariant>
#include <QString>
#include <QStack>
#include "DataBindingConfig.h"

class ReportDataModel;

class FormulaEngine : public QObject
{
    Q_OBJECT

public:
    explicit FormulaEngine(QObject* parent = nullptr);

    QVariant evaluate(const QString& formula, ReportDataModel* model, int currentRow, int currentCol);
    bool isFormula(const QString& text) const;

private:
    // 原有函数支持
    QVariant evaluateSum(const QString& range, ReportDataModel* model);
    QVariant evaluateMax(const QString& range, ReportDataModel* model);
    QVariant evaluateMin(const QString& range, ReportDataModel* model);

    // 新增：表达式计算支持
    QVariant evaluateExpression(const QString& expression, ReportDataModel* model);
    QString preprocessExpression(const QString& expression, ReportDataModel* model);
    QVariant calculateArithmetic(const QString& expression);

    // 操作符处理
    double applyOperator(double left, double right, QChar op);
    int getOperatorPrecedence(QChar op);
    bool isOperator(QChar c);

    // 单元格引用处理
    QPoint parseReference(const QString& ref) const;
    QPair<QPoint, QPoint> parseRange(const QString& range) const;
    QVariant getCellValue(const QString& cellRef, ReportDataModel* model);

    // 工具函数
    bool isValidCellReference(const QString& ref) const;
    QString columnToString(int col) const;
};

#endif // FORMULAENGINE_H