#pragma once
#ifndef DATABINDINGCONFIG_H
#define DATABINDINGCONFIG_H

#include <QDateTime>
#include <QString>
#include <QVariant>
#include <QFont>
#include <QSet>
#include <QColor>

// ===== 样式结构（从Cell.h移过来） =====
enum class RTBorderStyle {
    None = 0, Thin, Medium, Thick, Double, Dotted, Dashed
};

struct RTCellBorder {
    RTBorderStyle left = RTBorderStyle::None;
    RTBorderStyle right = RTBorderStyle::None;
    RTBorderStyle top = RTBorderStyle::None;
    RTBorderStyle bottom = RTBorderStyle::None;
    QColor leftColor = Qt::black;
    QColor rightColor = Qt::black;
    QColor topColor = Qt::black;
    QColor bottomColor = Qt::black;

    RTCellBorder() = default;
};

struct RTCellStyle {
    QFont font;
    QColor backgroundColor;
    QColor textColor = Qt::black;
    Qt::Alignment alignment = Qt::AlignLeft | Qt::AlignVCenter;
    RTCellBorder border;

    RTCellStyle() {
        font.setFamily("Arial");
        font.setPointSize(10);
        backgroundColor = Qt::white;
        font.setBold(false);
    }
};

// ===== 合并信息（从Cell.h移过来） =====
struct RTMergedRange {
    int startRow;
    int startCol;
    int endRow;
    int endCol;

    RTMergedRange() :startRow(-1), startCol(-1), endRow(-1), endCol(-1) {}
    RTMergedRange(int sRow, int sCol, int eRow, int eCol)
        :startRow(sRow), startCol(sCol), endRow(eRow), endCol(eCol) {
    }

    bool isValid() const {
        return startRow >= 0 && startCol >= 0 && endRow >= startRow && endCol >= startCol;
    }

    bool isMerged() const {
        return isValid() && (startRow != endRow || startCol != endCol);
    }

    bool contains(int row, int col) const {
        return isValid() && row >= startRow && row <= endRow && col >= startCol && col <= endCol;
    }

    int rowSpan() const { return isValid() ? (endRow - startRow + 1) : 1; }
    int colSpan() const { return isValid() ? (endCol - startCol + 1) : 1; }
};

// ===== 统一的单元格数据结构 =====
struct CellData {
    // 1. 基础数据
    QVariant value;
    QString formula;
    bool hasFormula = false;

    // 2. 样式
    RTCellStyle style;

    // 3. 合并信息
    RTMergedRange mergedRange;

    // 4. 数据绑定（单个绑定）
    bool isDataBinding = false;
    QString bindingKey;          // 单个绑定键，如 "##TEMP_01"

    CellData() = default;

    // 工具方法
    void setFormula(const QString& formulaText) {
        formula = formulaText;
        hasFormula = !formulaText.isEmpty() && formulaText.startsWith('=');
    }

    QString displayText() const {
        return value.toString();
    }

    QString editText() const {
        if (isDataBinding) {
            return bindingKey;  // 直接返回bindingKey
        }
        return hasFormula ? formula : value.toString();
    }

    bool isMergedMain() const {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    bool isMergedChild() const {
        return mergedRange.isValid() && isMergedMain() &&
            (mergedRange.startRow != 0 || mergedRange.startCol != 0);
    }
};

// ===== 时间范围配置 =====
struct TimeRangeConfig
{
    QDateTime startTime;
    QDateTime endTime;
    int intervalSeconds;

    TimeRangeConfig()
        : intervalSeconds(5)
    {
    }

    bool isValid() const {
        return startTime.isValid() && endTime.isValid() && intervalSeconds >= 0;  // 改为 >= 0，允许单点查询
    }

    bool isSinglePoint() const {
        return intervalSeconds == 0;  // 新增：判断是否为单点查询
    }
};

// ===== 全局配置 =====
struct GlobalDataConfig {
    TimeRangeConfig globalTimeRange;  // 全局时间范围
    QString defaultRTU;               // 默认RTU号

    GlobalDataConfig()
        : defaultRTU("AIRTU000100006")
    {
    }
};

// 报表列配置
struct ReportColumnConfig {
    QString displayName;     // 显示名称
    QString rtuId;           // RTU号
    int sourceRow;           // 配置文件中的行号

    ReportColumnConfig() : sourceRow(-1) {}
};

// 历史报表配置
struct HistoryReportConfig {
    QVector<ReportColumnConfig> columns;
    QString reportName;      // 报表名称
    QString configFilePath;  // 配置文件路径

    // 记录哪些列是数据列（只读）
    QSet<int> dataColumns;  // 存储数据列的列索引dataColumns 

    HistoryReportConfig() {}
};


#endif // DATABINDINGCONFIG_H