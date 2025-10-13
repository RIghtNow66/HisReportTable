#include "formulaengine.h"
#include "reportdatamodel.h"
#include <QRegularExpression>
#include <QStack>
#include <QQueue>

FormulaEngine::FormulaEngine(QObject* parent) : QObject(parent)
{
}

QVariant FormulaEngine::evaluate(const QString& formula, ReportDataModel* model, int currentRow, int currentCol)
{
    Q_UNUSED(currentRow)
        Q_UNUSED(currentCol)

        if (!isFormula(formula))
            return formula;

    QString expr = formula.mid(1).trimmed(); // 去掉 '='

    // 优先检查是否为函数调用
    QRegularExpression functionRegex(R"(^(SUM|MAX|MIN)\s*\([^)]+\)$)");
    QRegularExpressionMatch functionMatch = functionRegex.match(expr);

    if (functionMatch.hasMatch()) {
        // 处理函数
        if (expr.startsWith("SUM")) {
            QRegularExpression sumRegex(R"(SUM\(([A-Z]+\d+:[A-Z]+\d+)\))");
            QRegularExpressionMatch match = sumRegex.match(expr);
            if (match.hasMatch()) {
                return evaluateSum(match.captured(1), model);
            }
        }
        else if (expr.startsWith("MAX")) {
            QRegularExpression maxRegex(R"(MAX\(([A-Z]+\d+:[A-Z]+\d+)\))");
            QRegularExpressionMatch match = maxRegex.match(expr);
            if (match.hasMatch()) {
                return evaluateMax(match.captured(1), model);
            }
        }
        else if (expr.startsWith("MIN")) {
            QRegularExpression minRegex(R"(MIN\(([A-Z]+\d+:[A-Z]+\d+)\))");
            QRegularExpressionMatch match = minRegex.match(expr);
            if (match.hasMatch()) {
                return evaluateMin(match.captured(1), model);
            }
        }
    }

    // 如果不是函数，尝试作为表达式处理
    return evaluateExpression(expr, model);
}

bool FormulaEngine::isFormula(const QString& text) const
{
	return !text.isEmpty() && text.startsWith('=');
}

// 
QVariant FormulaEngine::evaluateSum(const QString& range, ReportDataModel* model)
{
    QPair<QPoint, QPoint> cellRange = parseRange(range);
    if (cellRange.first.x() == -1 || cellRange.second.x() == -1) {
        return QVariant(0.0);
    }

    double sum = 0.0;
    for (int row = cellRange.first.x(); row <= cellRange.second.x(); ++row) {
        for (int col = cellRange.first.y(); col <= cellRange.second.y(); ++col) {
            const CellData* cell = model->getCell(row, col);
            if (cell) {
                bool ok;
                double value = cell->value.toDouble(&ok);
                if (ok) {
                    sum += value;
                }
            }
        }
    }
    return QVariant(sum);
}

QVariant FormulaEngine::evaluateMax(const QString& range, ReportDataModel* model)
{
    QPair<QPoint, QPoint> cellRange = parseRange(range);
    if (cellRange.first.x() == -1 || cellRange.second.x() == -1) {
        return QVariant();
    }

    double maxValue = std::numeric_limits<double>::lowest();
    bool hasValue = false;

    for (int row = cellRange.first.x(); row <= cellRange.second.x(); ++row) {
        for (int col = cellRange.first.y(); col <= cellRange.second.y(); ++col) {
            const CellData* cell = model->getCell(row, col);
            if (cell) {
                bool ok;
                double value = cell->value.toDouble(&ok);
                if (ok) {
                    if (!hasValue || value > maxValue) {
                        maxValue = value;
                        hasValue = true;
                    }
                }
            }
        }
    }
    return hasValue ? QVariant(maxValue) : QVariant(0.0);
}

QVariant FormulaEngine::evaluateMin(const QString& range, ReportDataModel* model)
{
    QPair<QPoint, QPoint> cellRange = parseRange(range);
    if (cellRange.first.x() == -1 || cellRange.second.x() == -1) {
        return QVariant();
    }

    double minValue = std::numeric_limits<double>::max();
    bool hasValue = false;

    for (int row = cellRange.first.x(); row <= cellRange.second.x(); ++row) {
        for (int col = cellRange.first.y(); col <= cellRange.second.y(); ++col) {
            const CellData* cell = model->getCell(row, col);
            if (cell) {
                bool ok;
                double value = cell->value.toDouble(&ok);
                if (ok) {
                    if (!hasValue || value < minValue) {
                        minValue = value;
                        hasValue = true;
                    }
                }
            }
        }
    }
    return hasValue ? QVariant(minValue) : QVariant(0.0);
}


QVariant FormulaEngine::evaluateExpression(const QString& expression, ReportDataModel* model)
{
    // 1. 预处理：将单元格引用替换为实际值
    QString processedExpr = preprocessExpression(expression, model);

    // 2. 计算算术表达式
    return calculateArithmetic(processedExpr);
}

QString FormulaEngine::preprocessExpression(const QString& expression, ReportDataModel* model)
{
    QString result = expression;

    // 匹配单元格引用（如A1, B2, AA10等）
    QRegularExpression cellRefRegex(R"([A-Z]+\d+)");
    QRegularExpressionMatchIterator it = cellRefRegex.globalMatch(expression);

    // 从后往前替换，避免位置偏移
    QList<QRegularExpressionMatch> matches;
    while (it.hasNext()) {
        matches.append(it.next());
    }

    for (int i = matches.size() - 1; i >= 0; --i) {
        const QRegularExpressionMatch& match = matches[i];
        QString cellRef = match.captured(0);
        QVariant cellValue = getCellValue(cellRef, model);

        // 将单元格值转换为数字字符串
        bool ok;
        double numValue = cellValue.toDouble(&ok);
        if (!ok) {
            numValue = 0.0; // 如果不是数字，当作0处理
        }

        result.replace(match.capturedStart(), match.capturedLength(), QString::number(numValue));
    }

    return result;
}

QVariant FormulaEngine::calculateArithmetic(const QString& expression)
{
    // 使用调度场算法（Shunting Yard Algorithm）计算表达式
    QStack<double> numbers;
    QStack<QChar> operators;

    QString cleanExpr = expression;
    cleanExpr.remove(' '); // 移除空格

    int i = 0;
    while (i < cleanExpr.length()) {
        QChar c = cleanExpr[i];

        if (c.isDigit() || c == '.') {
            // 读取完整的数字
            QString numStr;
            while (i < cleanExpr.length() && (cleanExpr[i].isDigit() || cleanExpr[i] == '.')) {
                numStr += cleanExpr[i];
                i++;
            }
            numbers.push(numStr.toDouble());
            continue;
        }
        else if (c == '(') {
            operators.push(c);
        }
        else if (c == ')') {
            while (!operators.isEmpty() && operators.top() != '(') {
                QChar op = operators.pop();
                if (numbers.size() < 2) return QVariant("#ERROR!"); // 错误
                double right = numbers.pop();
                double left = numbers.pop();
                numbers.push(applyOperator(left, right, op));
            }
            if (!operators.isEmpty()) operators.pop(); // 移除 '('
        }
        else if (isOperator(c)) {
            while (!operators.isEmpty() &&
                operators.top() != '(' &&
                getOperatorPrecedence(operators.top()) >= getOperatorPrecedence(c)) {
                QChar op = operators.pop();
                if (numbers.size() < 2) return QVariant("#ERROR!"); // 错误
                double right = numbers.pop();
                double left = numbers.pop();
                numbers.push(applyOperator(left, right, op));
            }
            operators.push(c);
        }

        i++;
    }

    // 处理剩余的操作符
    while (!operators.isEmpty()) {
        QChar op = operators.pop();
        if (numbers.size() < 2) return QVariant("#ERROR!"); // 错误
        double right = numbers.pop();
        double left = numbers.pop();
        numbers.push(applyOperator(left, right, op));
    }

    if (numbers.size() == 1) {
        return numbers.top();
    }

    return QVariant("#ERROR!"); // 表达式错误
}

QVariant FormulaEngine::getCellValue(const QString& cellRef, ReportDataModel* model)
{
    QPoint pos = parseReference(cellRef);
    if (pos.x() == -1 || pos.y() == -1) {
        return QVariant(0.0);
    }

    const CellData* cell = model->getCell(pos.x(), pos.y());
    if (!cell) {
        return QVariant(0.0);
    }

    return cell->value;
}

double FormulaEngine::applyOperator(double left, double right, QChar op)
{
    switch (op.unicode()) {
    case '+': return left + right;
    case '-': return left - right;
    case '*': return left * right;
    case '/': return (right != 0) ? left / right : 0;
    default: return 0;
    }
}

int FormulaEngine::getOperatorPrecedence(QChar op)
{
    switch (op.unicode()) {
    case '+': case '-': return 1;
    case '*': case '/': return 2;
    default: return 0;
    }
}

bool FormulaEngine::isOperator(QChar c)
{
    return c == '+' || c == '-' || c == '*' || c == '/';
}

QPoint FormulaEngine::parseReference(const QString& ref) const
{
    if (!isValidCellReference(ref)) {
        return QPoint(-1, -1);
    }

    // 分离字母和数字部分
    QRegularExpression regex(R"(([A-Z]+)(\d+))");
    QRegularExpressionMatch match = regex.match(ref);

    if (!match.hasMatch()) {
        return QPoint(-1, -1);
    }

    QString colStr = match.captured(1);
    int rowNum = match.captured(2).toInt();

    // 转换列名为列索引
    int col = 0;
    for (int i = 0; i < colStr.length(); ++i) {
        col = col * 26 + (colStr[i].unicode() - 'A' + 1);
    }
    col -= 1; // 转换为0基索引

    return QPoint(rowNum - 1, col); // 转换为0基索引
}

// 解析范围字符串，如 "A1:B3"
QPair<QPoint, QPoint> FormulaEngine::parseRange(const QString& range) const
{
    QStringList parts = range.split(':');
    if (parts.size() != 2) {
        return qMakePair(QPoint(-1, -1), QPoint(-1, -1));
    }

    QPoint start = parseReference(parts[0].trimmed());
    QPoint end = parseReference(parts[1].trimmed());

    return qMakePair(start, end);
}

bool FormulaEngine::isValidCellReference(const QString& ref) const
{
    QRegularExpression regex(R"(^[A-Z]+\d+$)");
    return regex.match(ref).hasMatch();
}

QString FormulaEngine::columnToString(int col) const
{
    QString result;
    while (col >= 0) {
        result.prepend(QChar('A' + (col % 26)));
        col = col / 26 - 1;
    }
    return result;
}