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
    // ===== 基础字段 =====
    QVariant value;                    // 显示值
    QString formula;                   // 公式（如 "=SUM(A1:A10)"）
    bool hasFormula;                   // 是否有公式
    RTCellStyle style;                 // 样式
    RTMergedRange mergedRange;         // 合并信息

    // ===== 数据绑定字段（旧实时模式，保留兼容）=====
    bool isDataBinding;                // 是否是数据绑定（##RTU号）
    QString bindingKey;                // 绑定键（如 "##RTU001"）

    // ===== 日报标记字段（新增）=====
    enum CellType {
        NormalCell,         // 普通单元格
        DateMarker,         // #Date 标记
        TimeMarker,         // #t# 标记
        DataMarker          // #d# 标记（需要查询）
    };

    CellType cellType;                 // 单元格类型
    QString originalMarker;            // 原始标记文本（如 "#d#AIRTU034700019"）
    QString queryPath;                 // 查询地址（如 "AIRTU034700019@2024-01-01 00:00:00~..."）
    QString rtuId;                     // 提取的RTU号（如 "AIRTU034700019"）
    bool queryExecuted;                // 是否已执行查询
    bool querySuccess;                 // 查询是否成功

    // ===== 构造函数 =====
    CellData()
        : hasFormula(false)
        , isDataBinding(false)
        , cellType(NormalCell)
        , queryExecuted(false)
        , querySuccess(false)
    {
    }

    // ===== 工具函数 =====

    // 是否需要查询（日报模式）
    bool needsQuery() const {
        return cellType == DataMarker && !queryExecuted;
    }

    // 是否是标记单元格
    bool isMarker() const {
        return cellType != NormalCell;
    }

    // 是否是合并单元格的主单元格
    bool isMergedMain() const {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    // 设置公式
    void setFormula(const QString& formulaText) {
        formula = formulaText.startsWith('=') ? formulaText.mid(1) : formulaText;
        hasFormula = true;
    }

    // 获取显示文本
    QString displayText() const {
        if (hasFormula) {
            return value.toString();  // 公式显示计算结果
        }
        return value.toString();
    }

    // 获取编辑文本
    QString editText() const {
        if (hasFormula) {
            return "=" + formula;  // 编辑时显示公式
        }
        if (cellType == DataMarker && !originalMarker.isEmpty()) {
            return originalMarker;  // 日报标记编辑时显示原始标记
        }
        return value.toString();
    }
};

// ===== 时间范围配置（保留用于可能的扩展）=====
struct TimeRangeConfig {
    QDateTime startTime;
    QDateTime endTime;
    int intervalSeconds;

    TimeRangeConfig()
        : intervalSeconds(0)
    {
    }

    bool isValid() const {
        return startTime.isValid() && endTime.isValid() && intervalSeconds >= 0;
    }
};

// ===== 全局配置（保留用于可能的扩展）=====
struct GlobalDataConfig {
    TimeRangeConfig globalTimeRange;
};

#endif // DATABINDINGCONFIG_H