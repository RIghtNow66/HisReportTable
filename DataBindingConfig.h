#pragma once
#ifndef DATABINDINGCONFIG_H
#define DATABINDINGCONFIG_H

#include <QDateTime>
#include <QString>
#include <QVariant>
#include <QFont>
#include <QSet>
#include <QColor>

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

// ===== 单元格数据（核心结构）=====
struct CellData {
    // ===== 单元格类型枚举 =====
    enum CellType {
        NormalCell,         // 普通单元格
        DateMarker,         // #Date 标记
        TimeMarker,         // #t# 标记
        DataMarker          // #d# 标记（需要查询）
    };

    // ========================================
    // ===== 显示层（用户看到的） =====
    // ========================================
    QVariant displayValue;              // 显示值："2025年1月1日", "123.45", 公式计算结果

    // ========================================
    // ===== 标记层（程序逻辑用的） =====
    // ========================================
    QString markerText;                 // 原始标记："#Date:2025-01-01", "#t#0:00", "#d#RTU001"
    CellType cellType;                  // 单元格类型
    QString rtuId;                      // 数据标记的RTU号（从markerText解析）

    // ========================================
    // ===== 公式相关 =====
    // ========================================
    bool hasFormula;                    // 是否有公式
    QString formula;                    // 公式文本："=A1+B1"
    bool formulaCalculated;             // 公式是否已计算

    // ========================================
    // ===== 查询状态 =====
    // ========================================
    bool queryExecuted;                 // 是否已执行查询
    bool querySuccess;                  // 查询是否成功

    // ========================================
    // ===== 样式 =====
    // ========================================
    RTCellStyle style;                  // 样式
    RTMergedRange mergedRange;          // 合并信息

    // ========================================
    // ===== 兼容性字段（保留用于向后兼容，后续可删除） =====
    // ========================================
    QVariant value;                     // 已废弃，使用 displayValue 代替
    QString originalMarker;             // 已废弃，使用 markerText 代替
    bool isDataBinding;                 // 已废弃（旧实时模式）
    QString bindingKey;                 // 已废弃（旧实时模式）
    QString queryPath;                  // 已废弃

    // ===== 构造函数 =====
    CellData()
        : displayValue()
        , markerText()
        , cellType(NormalCell)
        , rtuId()
        , hasFormula(false)
        , formula()
        , formulaCalculated(false)
        , queryExecuted(false)
        , querySuccess(false)
        , style()
        , mergedRange()
        , value()                       // 兼容性字段
        , originalMarker()              // 兼容性字段
        , isDataBinding(false)          // 兼容性字段
        , bindingKey()                  // 兼容性字段
        , queryPath()                   // 兼容性字段
    {
    }

    // ========================================
    // ===== 核心判断方法 =====
    // ========================================

    // 是否是标记单元格
    bool isMarker() const {
        return cellType != NormalCell;
    }

    // 是否需要查询（日报模式）
    bool needsQuery() const {
        return cellType == DataMarker && !queryExecuted;
    }

    // 是否是合并单元格的主单元格
    bool isMergedMain() const {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    // ========================================
    // ===== 文本获取方法（用于不同场景） =====
    // ========================================

    /**
     * 获取显示文本（Model的DisplayRole用）
     * 优先级：公式计算结果 > 公式文本 > 显示值
     */
    QString displayText() const {
        if (hasFormula && formulaCalculated) {
            return displayValue.toString();  // 显示公式计算结果
        }
        if (hasFormula && !formulaCalculated) {
            return formula;  // 显示公式文本
        }
        return displayValue.toString();  // 显示普通值或转换后的标记
    }

    /**
     * 获取编辑文本（Model的EditRole用）
     * 优先级：公式 > 标记 > 显示值
     */
    QString editText() const {
        if (hasFormula) {
            return formula;  // 编辑时显示公式
        }
        if (!markerText.isEmpty()) {
            return markerText;  // 编辑时显示原始标记
        }
        // 兼容旧数据
        if (!originalMarker.isEmpty()) {
            return originalMarker;
        }
        return displayValue.toString();  // 编辑普通值
    }

    /**
     * 获取用于扫描的文本（解析器用）
     * 优先级：标记文本 > 兼容字段 > 显示值
     */
    QString scanText() const {
        // 优先使用标记文本
        if (!markerText.isEmpty()) {
            return markerText;
        }
        // 兼容旧数据：尝试从 originalMarker 读取
        if (!originalMarker.isEmpty()) {
            return originalMarker;
        }
        // 兼容旧数据：尝试从 value 读取
        if (!value.isNull() && !value.toString().isEmpty()) {
            return value.toString();
        }
        // 最后降级到 displayValue
        return displayValue.toString();
    }

    // ========================================
    // ===== 公式操作方法 =====
    // ========================================

    /**
     * 设置公式
     */
    void setFormula(const QString& formulaText) {
        formula = formulaText;
        hasFormula = true;
        formulaCalculated = false;
        markerText.clear();  // 公式不是标记
        cellType = NormalCell;
    }
};

struct ReportColumnConfig {
    QString displayName;  // 显示名称
    QString rtuId;        // RTU号
    int sourceRow;        // 源文件行号（可选）

    ReportColumnConfig() : sourceRow(-1) {}
};

struct HistoryReportConfig {
    QString reportName;                    // 报表名称
    QString configFilePath;                // 配置文件路径
    QList<ReportColumnConfig> columns;     // 列配置列表
    QSet<int> dataColumns;                 // 数据列索引集合（用于标记只读列）
};

struct TimeRangeConfig {
    QDateTime startTime;
    QDateTime endTime;
    int intervalSeconds;

    bool isValid() const {
        return startTime.isValid() &&
            endTime.isValid() &&
            intervalSeconds >= 0 &&
            startTime <= endTime;
    }
};

// ===== 全局配置（保留用于可能的扩展）=====
struct GlobalDataConfig {
    TimeRangeConfig globalTimeRange;
};

#endif // DATABINDINGCONFIG_H