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

// ===== �ϲ���Ϣ����Cell.h�ƹ����� =====
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

// ===== ��Ԫ�����ݣ����Ľṹ��=====
struct CellData {
    // ===== �����ֶ� =====
    QVariant value;                    // ��ʾֵ
    QString formula;                   // ��ʽ���� "=SUM(A1:A10)"��
    bool hasFormula;                   // �Ƿ��й�ʽ
    RTCellStyle style;                 // ��ʽ
    RTMergedRange mergedRange;         // �ϲ���Ϣ

    // ===== ���ݰ��ֶΣ���ʵʱģʽ���������ݣ�=====
    bool isDataBinding;                // �Ƿ������ݰ󶨣�##RTU�ţ�
    QString bindingKey;                // �󶨼����� "##RTU001"��

    // ===== �ձ�����ֶΣ�������=====
    enum CellType {
        NormalCell,         // ��ͨ��Ԫ��
        DateMarker,         // #Date ���
        TimeMarker,         // #t# ���
        DataMarker          // #d# ��ǣ���Ҫ��ѯ��
    };

    CellType cellType;                 // ��Ԫ������
    QString originalMarker;            // ԭʼ����ı����� "#d#AIRTU034700019"��
    QString queryPath;                 // ��ѯ��ַ���� "AIRTU034700019@2024-01-01 00:00:00~..."��
    QString rtuId;                     // ��ȡ��RTU�ţ��� "AIRTU034700019"��
    bool queryExecuted;                // �Ƿ���ִ�в�ѯ
    bool querySuccess;                 // ��ѯ�Ƿ�ɹ�

    // ===== ���캯�� =====
    CellData()
        : hasFormula(false)
        , isDataBinding(false)
        , cellType(NormalCell)
        , queryExecuted(false)
        , querySuccess(false)
    {
    }

    // ===== ���ߺ��� =====

    // �Ƿ���Ҫ��ѯ���ձ�ģʽ��
    bool needsQuery() const {
        return cellType == DataMarker && !queryExecuted;
    }

    // �Ƿ��Ǳ�ǵ�Ԫ��
    bool isMarker() const {
        return cellType != NormalCell;
    }

    // �Ƿ��Ǻϲ���Ԫ�������Ԫ��
    bool isMergedMain() const {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    // ���ù�ʽ
    void setFormula(const QString& formulaText) {
        formula = formulaText.startsWith('=') ? formulaText.mid(1) : formulaText;
        hasFormula = true;
    }

    // ��ȡ��ʾ�ı�
    QString displayText() const {
        if (hasFormula) {
            return value.toString();  // ��ʽ��ʾ������
        }
        return value.toString();
    }

    // ��ȡ�༭�ı�
    QString editText() const {
        if (hasFormula) {
            return "=" + formula;  // �༭ʱ��ʾ��ʽ
        }
        if (cellType == DataMarker && !originalMarker.isEmpty()) {
            return originalMarker;  // �ձ���Ǳ༭ʱ��ʾԭʼ���
        }
        return value.toString();
    }
};

// ===== ʱ�䷶Χ���ã��������ڿ��ܵ���չ��=====
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

// ===== ȫ�����ã��������ڿ��ܵ���չ��=====
struct GlobalDataConfig {
    TimeRangeConfig globalTimeRange;
};

#endif // DATABINDINGCONFIG_H