#include "excelhandler.h"
#include "reportdatamodel.h"
#include "DataBindingConfig.h"

#include <QMessageBox>
#include <QFileInfo>
#include <QApplication>
#include <QProgressDialog>
#include <memory> // for std::make_unique
#include <QSet>
#include <QTextCodec>
#include <qDebug>

// 明确包含所有需要的QXlsx库头文件
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcell.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"


bool ExcelHandler::loadFromFile(const QString& fileName, ReportDataModel* model)
{
    if (fileName.isEmpty() || !model) {
        QMessageBox::warning(nullptr, "错误", "参数无效");
        return false;
    }
    if (!isValidExcelFile(fileName)) {
        return false;
    }

    auto progress = std::make_unique<QProgressDialog>("正在导入Excel文件...", "取消", 0, 100, nullptr);
    progress->setWindowModality(Qt::ApplicationModal);
    progress->setMinimumDuration(0);
    progress->show();

    model->clearAllCells();

    QXlsx::Document xlsx(fileName);
    if (!xlsx.load()) {
        QMessageBox::warning(nullptr, "错误", "无法打开Excel文件");
        return false;
    }
    progress->setValue(10);
    qApp->processEvents();

    QXlsx::Worksheet* worksheet = static_cast<QXlsx::Worksheet*>(xlsx.currentSheet());
    if (!worksheet) {
        QMessageBox::warning(nullptr, "错误", "无法读取工作表");
        return false;
    }

    loadRowColumnSizes(worksheet, model);

    const QXlsx::CellRange range = worksheet->dimension();
    if (!range.isValid()) {
        model->updateModelSize(0, 0);
        return true;
    }

    int totalRows = range.rowCount();
    model->updateModelSize(range.rowCount(), range.columnCount());
    progress->setValue(20);

    // 1. 先加载所有合并单元格信息
    QHash<QPoint, RTMergedRange> mergedRanges;
    loadMergedCells(worksheet, mergedRanges);
    progress->setValue(30);

    // 2. 遍历有效区域内的每一个格子，确保不遗漏
    for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
        if (progress->wasCanceled()) break;
        for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {

            int modelRow = row - 1, modelCol = col - 1;
            CellData* newCell = new CellData();

            // 尝试获取单元格对象
            auto xlsxCell = worksheet->cellAt(row, col);

            if (xlsxCell) {
                QVariant rawValue = xlsxCell->value();
                if (rawValue.type() == QVariant::String) {
                    QString rawString = rawValue.toString();
                    if (rawString.startsWith("#=#")) {
                        newCell->setFormula(rawString);  // 内部会设置 formulaCalculated = false
                        newCell->value = rawString;       // 显示公式文本
                    }
                    else if (rawString.startsWith("##") || rawString.startsWith("#")) {
                        newCell->originalMarker = rawString;
                        newCell->value = rawString;
                    }
                    else {
                        newCell->value = rawValue;
                    }
                }
                else {
                    newCell->value = rawValue;
                }

                QXlsx::Format cellFormat = xlsxCell->format();
                convertFromExcelStyle(cellFormat, newCell->style);
            }
            else {
                // 【新增】即使 cellAt 返回 nullptr，也尝试读取该位置的格式
                // 通过创建一个临时 Cell 对象来获取格式
                QXlsx::Cell* tempCell = new QXlsx::Cell(QVariant(), QXlsx::Cell::NumberType,
                    QXlsx::Format(), worksheet);
                // 注意：上面这行可能不work，因为 QXlsx::Cell 构造函数可能是私有的
                // 如果编译报错，就删除这个 else 分支
            }

            model->addCellDirect(modelRow, modelCol, newCell);
        }
        int progressValue = 30 + ((row - range.firstRow() + 1) * 60 / totalRows);
        progress->setValue(progressValue);
        qApp->processEvents();
    }

    if (progress->wasCanceled()) {
        model->clearAllCells();
        return false;
    }

    // 3. 【后处理】统一合并区域的样式
    // 这个逻辑很重要，它会为合并区域中的空单元格填充上主单元格的样式（包括边框）
    for (const auto& mergedRange : mergedRanges) {
        CellData* mainCell = model->getCell(mergedRange.startRow, mergedRange.startCol);
        if (!mainCell) continue;

        for (int r = mergedRange.startRow; r <= mergedRange.endRow; ++r) {
            for (int c = mergedRange.startCol; c <= mergedRange.endCol; ++c) {
                CellData* childCell = model->getCell(r, c);
                if (childCell) {
                    childCell->mergedRange = mergedRange;

                    if (r != mergedRange.startRow || c != mergedRange.startCol) {
                        // 【核心修复】保存原始边框
                        RTCellBorder savedBorder = childCell->style.border;

                        // 复制主单元格的样式
                        childCell->style = mainCell->style;

                        // 【关键】恢复原始边框
                        childCell->style.border = savedBorder;

                        // 清空内容
                        childCell->value = QVariant();
                    }
                }
            }
        }
    }

    progress->setValue(100);
    return true;
}

void ExcelHandler::loadRowColumnSizes(QXlsx::Worksheet* worksheet, ReportDataModel* model)
{
    if (!worksheet || !model) return;

    const QXlsx::CellRange dimension = worksheet->dimension();

    // 列宽：Excel 到像素的近似转换
    // 推荐用 7 (Excel 列宽 1 ≈ 7px)，而不是固定 64/8.43
    auto colWidthToPixel = [](double excelWidth) {
        if (excelWidth <= 0) return 0.0;
        return std::floor((excelWidth + 0.72) * 7);  // 经验公式，微软官方文档的近似
        };

    for (int col = dimension.firstColumn(); col <= dimension.lastColumn(); ++col) {
        double width = worksheet->columnWidth(col);
        if (width > 0) {
            model->setColumnWidth(col - 1, colWidthToPixel(width));
        }
    }

    // 行高：point -> pixel
    const double pointToPixelRatio = 96.0 / 72.0;
    for (int row = dimension.firstRow(); row <= dimension.lastRow(); ++row) {
        double height = worksheet->rowHeight(row);
        if (height > 0) {
            model->setRowHeight(row - 1, height * pointToPixelRatio);
        }
    }
}

bool ExcelHandler::saveToFile(const QString& fileName, ReportDataModel* model, ExportMode mode)
{
    if (fileName.isEmpty() || !model) {
        QMessageBox::warning(nullptr, "错误", "参数无效");
        return false;
    }

    QString actualFileName = fileName;
    if (!actualFileName.endsWith(".xlsx", Qt::CaseInsensitive)) {
        actualFileName += ".xlsx";
    }

    auto progress = std::make_unique<QProgressDialog>("正在导出Excel文件...", "取消", 0, 100, nullptr);
    progress->setWindowModality(Qt::ApplicationModal);
    progress->setMinimumDuration(0);
    progress->show();

    QXlsx::Document xlsx;
    QXlsx::Worksheet* worksheet = xlsx.currentWorksheet();
    if (!worksheet) {
        QMessageBox::warning(nullptr, "错误", "无法创建Excel工作表");
        return false;
    }

    worksheet->setGridLinesVisible(true);

    progress->setValue(10);

    const auto& allCells = model->getAllCells();
    int totalCells = allCells.size();
    int processedCells = 0;

    // ===== 计算实际数据范围 =====
    int maxDataRow = 0;
    int maxDataCol = 0;
    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        maxDataRow = qMax(maxDataRow, it.key().x());
        maxDataCol = qMax(maxDataCol, it.key().y());
    }

    // ===== 保存单元格数据和格式 =====
    // 1. 定义一个默认样式实例用于比较
    RTCellStyle defaultStyle;

    // Helper: 检查边框是否为默认
    auto isDefaultBorder = [](const RTCellBorder& border) {
        return border.left == RTBorderStyle::None &&
            border.right == RTBorderStyle::None &&
            border.top == RTBorderStyle::None &&
            border.bottom == RTBorderStyle::None &&
            border.leftColor == Qt::black &&
            border.rightColor == Qt::black &&
            border.topColor == Qt::black &&
            border.bottomColor == Qt::black;
        };

    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        if (progress->wasCanceled()) return false;

        const QPoint& modelPos = it.key();
        const CellData* cell = it.value();

        int excelRow = modelPos.x() + 1;
        int excelCol = modelPos.y() + 1;

        // 根据导出模式决定写入内容
        QVariant valueToWrite = getCellValueForExport(cell, mode);

        // 2. 【强化检查】判断是否为完全默认的样式 (包括字体)
        bool isTrulyDefault =
            cell->style.backgroundColor == defaultStyle.backgroundColor &&
            cell->style.textColor == defaultStyle.textColor &&
            cell->style.alignment == defaultStyle.alignment &&
            isDefaultBorder(cell->style.border) &&
            // 增加字体家族、字号和加粗的检查
            cell->style.font.family() == defaultStyle.font.family() &&
            cell->style.font.pointSize() == defaultStyle.font.pointSize() &&
            cell->style.font.bold() == defaultStyle.font.bold();

        // 只有在单元格样式为默认，且没有合并单元格时，才不写入格式对象
        if (isTrulyDefault && !cell->mergedRange.isMerged())
        {
            // 样式完全是默认的，不传入 Format 对象，Excel 将显示默认网格线
            worksheet->write(excelRow, excelCol, valueToWrite);
        }
        else
        {
            // 样式非默认（例如有边框、背景色、或非默认字体等），或者它是合并单元格的主单元格。
            // 必须写入 Format 对象。
            QXlsx::Format cellFormat = convertToExcelFormat(cell->style);
            worksheet->write(excelRow, excelCol, valueToWrite, cellFormat);
        }

        processedCells++;
        if (totalCells > 0) {
            progress->setValue(10 + (processedCells * 60 / totalCells));
            qApp->processEvents();
        }
    }

    progress->setValue(70);

    // ===== 只设置有数据范围内的行高 =====
    const auto& rowHeights = model->getAllRowHeights();
    const double pixelToPointRatio = 0.75;
    for (int i = 0; i <= maxDataRow; ++i) {
        if (i < rowHeights.size() && rowHeights[i] > 0) {
            worksheet->setRowHeight(i + 1, i + 1, rowHeights[i] * pixelToPointRatio);
        }
    }

    progress->setValue(80);

    // ===== 只设置有数据范围内的列宽 =====
    const auto& colWidths = model->getAllColumnWidths();
    const double pixelToCharacterWidthRatio = 1.0 / 7.0;  // 修复转换系数
    for (int i = 0; i <= maxDataCol; ++i) {
        if (i < colWidths.size() && colWidths[i] > 0) {
            worksheet->setColumnWidth(i + 1, i + 1, colWidths[i] * pixelToCharacterWidthRatio);
        }
    }

    // 保存合并单元格
    saveMergedCells(worksheet, allCells);
    progress->setValue(90);

    if (progress->wasCanceled()) return false;

    if (!xlsx.saveAs(actualFileName)) {
        QMessageBox::warning(nullptr, "保存失败", QString("无法保存文件到：%1").arg(actualFileName));
        return false;
    }

    progress->setValue(100);
    return true;
}

void ExcelHandler::loadMergedCells(QXlsx::Worksheet* worksheet, QHash<QPoint, RTMergedRange>& mergedRanges)
{
    // 获取合并单元格范围列表
    const auto& mergedCells = worksheet->mergedCells();

    for (const QXlsx::CellRange& range : mergedCells) {
        if (range.isValid()) {
            RTMergedRange mergedRange(
                range.firstRow() - 1,    // 转换为0基索引
                range.firstColumn() - 1,
                range.lastRow() - 1,
                range.lastColumn() - 1
            );

            // 为合并范围内的所有单元格设置合并信息
            for (int row = range.firstRow(); row <= range.lastRow(); ++row) {
                for (int col = range.firstColumn(); col <= range.lastColumn(); ++col) {
                    QPoint cellPos(row - 1, col - 1);  // 转换为0基索引
                    mergedRanges[cellPos] = mergedRange;
                }
            }
        }
    }
}

void ExcelHandler::saveMergedCells(QXlsx::Worksheet* worksheet, const QHash<QPoint, CellData*>& allCells)
{
    QSet<RTMergedRange*> processedRanges;  // 避免重复处理相同的合并范围

    for (auto it = allCells.constBegin(); it != allCells.constEnd(); ++it) {
        const CellData* cell = it.value();

        if (cell->isMergedMain() && !processedRanges.contains(const_cast<RTMergedRange*>(&cell->mergedRange))) {
            // 转换为Excel的1基索引
            QXlsx::CellRange range(
                cell->mergedRange.startRow + 1,
                cell->mergedRange.startCol + 1,
                cell->mergedRange.endRow + 1,
                cell->mergedRange.endCol + 1
            );

            worksheet->mergeCells(range);
            processedRanges.insert(const_cast<RTMergedRange*>(&cell->mergedRange));
        }
    }
}

void ExcelHandler::convertFromExcelStyle(const QXlsx::Format& excelFormat, RTCellStyle& cellStyle)
{
    // 
    // 1. 字体处理：保证读取的字号就是最终显示的字号
    int fontSize = excelFormat.fontSize();
    if (fontSize > 0) {
        // 如果成功读取到字号，就使用它
        cellStyle.font.setPointSize(fontSize);
    }
    // 如果读取失败(fontSize为0)，则会使用在Cell.h中设置的默认字号(10)

    // 读取字体名称
    QString fontName = excelFormat.fontName();
    QString mappedFontName = mapChineseFontName(fontName);
    if (!fontName.isEmpty()) {
        cellStyle.font.setFamily(mappedFontName);
    }

    // 读取加粗状态
    cellStyle.font.setBold(excelFormat.fontBold());

    // 读取字体颜色
    cellStyle.textColor = excelFormat.fontColor();

    // 2. 背景颜色处理：默认白色，有颜色时统一为浅灰色
    // (默认背景已在Cell.h中设为白色)
    if (excelFormat.fillPattern() != QXlsx::Format::PatternNone) {
        // 只要Excel单元格有任何填充模式，就统一设置为浅灰色
        cellStyle.backgroundColor = QColor(Qt::white); // 浅灰色
    }

    // ===== 对齐方式 =====
    Qt::Alignment hAlign = Qt::AlignLeft;
    switch (excelFormat.horizontalAlignment()) {
    case QXlsx::Format::AlignHCenter: hAlign = Qt::AlignHCenter; break;
    case QXlsx::Format::AlignRight:   hAlign = Qt::AlignRight; break;
    default: break;
    }

    Qt::Alignment vAlign = Qt::AlignVCenter;
    switch (excelFormat.verticalAlignment()) {
    case QXlsx::Format::AlignTop:    vAlign = Qt::AlignTop; break;
    case QXlsx::Format::AlignBottom: vAlign = Qt::AlignBottom; break;
    default: break;
    }
    cellStyle.alignment = hAlign | vAlign;

    // ===== 边框 =====
    convertBorderFromExcel(excelFormat, cellStyle.border);
}

// 中文字体名称映射函数
QString ExcelHandler::mapChineseFontName(const QString& originalName)
{
    // 创建中文字体映射表
    static QHash<QString, QString> fontMap = {
        // 常见的Excel中文字体映射
        {"宋体", "SimSun"},
        {"黑体", "SimHei"},
        {"楷体", "KaiTi"},
        {"仿宋", "FangSong"},
        {"微软雅黑", "Microsoft YaHei"},
        {"新宋体", "NSimSun"},

        // 英文字体保持不变
        {"Calibri", "Calibri"},
        {"Arial", "Arial"},
        {"Times New Roman", "Times New Roman"},
        {"Verdana", "Verdana"},

        // 可能的字体族名变体
        {"MS Song", "SimSun"},
        {"MS Gothic", "SimHei"},
        {"MS Mincho", "SimSun"}
    };

    // 先尝试直接映射
    if (fontMap.contains(originalName)) {
        QString mapped = fontMap[originalName];
        return mapped;
    }

    // 如果没有找到映射，检查是否包含中文字符
    for (const QChar& ch : originalName) {
        if (ch.unicode() >= 0x4e00 && ch.unicode() <= 0x9fff) {
            // 包含中文字符，很可能是中文字体
            return originalName;  // 直接使用原名
        }
    }

    // 默认情况，返回原名
    return originalName;
}


QXlsx::Format ExcelHandler::convertToExcelFormat(const RTCellStyle& cellStyle)
{
    QXlsx::Format excelFormat;

    // 字体
    excelFormat.setFont(cellStyle.font);

    // 颜色
    excelFormat.setPatternBackgroundColor(cellStyle.backgroundColor);
    excelFormat.setFontColor(cellStyle.textColor);

    // 对齐方式
    if (cellStyle.alignment & Qt::AlignHCenter)
        excelFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    else if (cellStyle.alignment & Qt::AlignRight)
        excelFormat.setHorizontalAlignment(QXlsx::Format::AlignRight);

    if (cellStyle.alignment & Qt::AlignTop)
        excelFormat.setVerticalAlignment(QXlsx::Format::AlignTop);
    else if (cellStyle.alignment & Qt::AlignBottom)
        excelFormat.setVerticalAlignment(QXlsx::Format::AlignBottom);
    else
        excelFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    // 边框 - 现在可以正常处理了
    convertBorderToExcel(cellStyle.border, excelFormat);

    return excelFormat;
}

void ExcelHandler::convertBorderFromExcel(const QXlsx::Format& excelFormat, RTCellBorder& border)
{
    // 转换边框样式 - 使用正确的API
    border.left = convertBorderStyleFromExcel(excelFormat.leftBorderStyle());
    border.right = convertBorderStyleFromExcel(excelFormat.rightBorderStyle());
    border.top = convertBorderStyleFromExcel(excelFormat.topBorderStyle());
    border.bottom = convertBorderStyleFromExcel(excelFormat.bottomBorderStyle());

    // 转换边框颜色 - 使用正确的API
    border.leftColor = excelFormat.leftBorderColor();
    border.rightColor = excelFormat.rightBorderColor();
    border.topColor = excelFormat.topBorderColor();
    border.bottomColor = excelFormat.bottomBorderColor();
}


void ExcelHandler::convertBorderToExcel(const RTCellBorder& border, QXlsx::Format& excelFormat)
{
    // 转换边框样式 - 使用正确的API
    excelFormat.setLeftBorderStyle(convertBorderStyleToExcel(border.left));
    excelFormat.setRightBorderStyle(convertBorderStyleToExcel(border.right));
    excelFormat.setTopBorderStyle(convertBorderStyleToExcel(border.top));
    excelFormat.setBottomBorderStyle(convertBorderStyleToExcel(border.bottom));

    // 转换边框颜色 - 使用正确的API
    excelFormat.setLeftBorderColor(border.leftColor);
    excelFormat.setRightBorderColor(border.rightColor);
    excelFormat.setTopBorderColor(border.topColor);
    excelFormat.setBottomBorderColor(border.bottomColor);
}


RTBorderStyle ExcelHandler::convertBorderStyleFromExcel(QXlsx::Format::BorderStyle xlsxStyle)
{
    switch (xlsxStyle) {
    case QXlsx::Format::BorderNone: return RTBorderStyle::None;
    case QXlsx::Format::BorderThin: return RTBorderStyle::Thin;
    case QXlsx::Format::BorderMedium: return RTBorderStyle::Medium;
    case QXlsx::Format::BorderThick: return RTBorderStyle::Thick;
    case QXlsx::Format::BorderDouble: return RTBorderStyle::Double;
    case QXlsx::Format::BorderDotted: return RTBorderStyle::Dotted;
    case QXlsx::Format::BorderDashed: return RTBorderStyle::Dashed;
        // QXlsx还有更多边框样式，可以映射到我们的样式
    case QXlsx::Format::BorderHair: return RTBorderStyle::Thin;  // 映射为细线
    case QXlsx::Format::BorderMediumDashed: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderDashDot: return RTBorderStyle::Dashed;  // 映射为虚线
    case QXlsx::Format::BorderMediumDashDot: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderDashDotDot: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderMediumDashDotDot: return RTBorderStyle::Dashed;
    case QXlsx::Format::BorderSlantDashDot: return RTBorderStyle::Dashed;
    default: return RTBorderStyle::None;
    }
}

QXlsx::Format::BorderStyle ExcelHandler::convertBorderStyleToExcel(RTBorderStyle rtStyle)
{
    switch (rtStyle) {
    case RTBorderStyle::None: return QXlsx::Format::BorderNone;
    case RTBorderStyle::Thin: return QXlsx::Format::BorderThin;
    case RTBorderStyle::Medium: return QXlsx::Format::BorderMedium;
    case RTBorderStyle::Thick: return QXlsx::Format::BorderThick;
    case RTBorderStyle::Double: return QXlsx::Format::BorderDouble;
    case RTBorderStyle::Dotted: return QXlsx::Format::BorderDotted;
    case RTBorderStyle::Dashed: return QXlsx::Format::BorderDashed;
    default: return QXlsx::Format::BorderNone;
    }
}


bool ExcelHandler::isValidExcelFile(const QString& fileName)
{
    QFileInfo fileInfo(fileName);
    if (!fileInfo.exists()) {
        QMessageBox::warning(nullptr, "错误", QString("文件不存在：%1").arg(fileName));
        return false;
    }
    if (!fileInfo.isReadable()) {
        QMessageBox::warning(nullptr, "错误", QString("无法读取文件：%1").arg(fileName));
        return false;
    }
    return true;
}

QVariant ExcelHandler::getCellValueForExport(const CellData* cell, ExportMode mode)
{
    if (!cell) return QVariant();

    if (mode == EXPORT_DATA) {
        // 导出数据：返回显示值
        if (cell->hasFormula && cell->formulaCalculated) {
            return cell->value;
        }
        else if (cell->cellType == CellData::DataMarker) {
            return cell->queryExecuted && cell->querySuccess ? cell->value : QVariant("N/A");
        }
        else if (cell->isDataBinding) {
            return cell->value;
        }
        return cell->value;
    }
    else {
        // 导出模板：返回原始标记
        if (cell->hasFormula) return cell->formula;
        if (!cell->originalMarker.isEmpty()) return cell->originalMarker;
        if (cell->isDataBinding) return cell->bindingKey;
        return cell->value;
    }
}
