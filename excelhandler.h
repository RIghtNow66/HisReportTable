#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include <QString>
#include <QHash>
#include <QPoint>
#include <QFontInfo> 

#include "xlsxformat.h"

// 前向声明所有需要的类型
class ReportDataModel;
struct CellData;
struct RTCellStyle;
struct RTCellBorder;      // 添加这个声明
struct RTMergedRange;
enum class RTBorderStyle; // 添加这个声明

namespace QXlsx {
    class Worksheet;
}

class ExcelHandler
{
public:
    enum ExportMode {
        EXPORT_DATA,
        EXPORT_TEMPLATE
    };

    static bool loadFromFile(const QString& fileName, ReportDataModel* model);
    static bool saveToFile(const QString& fileName, ReportDataModel* model,
        ExportMode mode = EXPORT_DATA);
private:
    ExcelHandler() = delete;
    ~ExcelHandler() = delete;
    static void convertFromExcelStyle(const QXlsx::Format& excelFormat, RTCellStyle& cellStyle);
    static QXlsx::Format convertToExcelFormat(const RTCellStyle& cellStyle);
    static bool isValidExcelFile(const QString& fileName);

    static void loadMergedCells(QXlsx::Worksheet* worksheet, QHash<QPoint, RTMergedRange>& mergedRanges);
    static void saveMergedCells(QXlsx::Worksheet* worksheet, const QHash<QPoint, CellData*>& allCells);
    static void convertBorderFromExcel(const QXlsx::Format& excelFormat, RTCellBorder& border);
    static void convertBorderToExcel(const RTCellBorder& border, QXlsx::Format& excelFormat);
    static RTBorderStyle convertBorderStyleFromExcel(QXlsx::Format::BorderStyle xlsxStyle);
    static QXlsx::Format::BorderStyle convertBorderStyleToExcel(RTBorderStyle rtStyle);

    //void loadRowColumnSizes(QXlsx::Worksheet* worksheet, ReportDataModel* model);

    static QString mapChineseFontName(const QString& originalName);

    static void loadRowColumnSizes(QXlsx::Worksheet* worksheet, ReportDataModel* model);

    static QVariant getCellValueForExport(const CellData* cell, ExportMode mode);
};
#endif // EXCELHANDLER_H
