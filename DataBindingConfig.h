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
    // ===== ��Ԫ������ö�� =====
    enum CellType {
        NormalCell,         // ��ͨ��Ԫ��
        DateMarker,         // #Date ���
        TimeMarker,         // #t# ���
        DataMarker          // #d# ��ǣ���Ҫ��ѯ��
    };

    // ========================================
    // ===== ��ʾ�㣨�û������ģ� =====
    // ========================================
    QVariant displayValue;              // ��ʾֵ��"2025��1��1��", "123.45", ��ʽ������

    // ========================================
    // ===== ��ǲ㣨�����߼��õģ� =====
    // ========================================
    QString markerText;                 // ԭʼ��ǣ�"#Date:2025-01-01", "#t#0:00", "#d#RTU001"
    CellType cellType;                  // ��Ԫ������
    QString rtuId;                      // ���ݱ�ǵ�RTU�ţ���markerText������

    // ========================================
    // ===== ��ʽ��� =====
    // ========================================
    bool hasFormula;                    // �Ƿ��й�ʽ
    QString formula;                    // ��ʽ�ı���"=A1+B1"
    bool formulaCalculated;             // ��ʽ�Ƿ��Ѽ���

    // ========================================
    // ===== ��ѯ״̬ =====
    // ========================================
    bool queryExecuted;                 // �Ƿ���ִ�в�ѯ
    bool querySuccess;                  // ��ѯ�Ƿ�ɹ�

    // ========================================
    // ===== ��ʽ =====
    // ========================================
    RTCellStyle style;                  // ��ʽ
    RTMergedRange mergedRange;          // �ϲ���Ϣ

    // ========================================
    // ===== �������ֶΣ��������������ݣ�������ɾ���� =====
    // ========================================
    QVariant value;                     // �ѷ�����ʹ�� displayValue ����
    QString originalMarker;             // �ѷ�����ʹ�� markerText ����
    bool isDataBinding;                 // �ѷ�������ʵʱģʽ��
    QString bindingKey;                 // �ѷ�������ʵʱģʽ��
    QString queryPath;                  // �ѷ���

    // ===== ���캯�� =====
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
        , value()                       // �������ֶ�
        , originalMarker()              // �������ֶ�
        , isDataBinding(false)          // �������ֶ�
        , bindingKey()                  // �������ֶ�
        , queryPath()                   // �������ֶ�
    {
    }

    // ========================================
    // ===== �����жϷ��� =====
    // ========================================

    // �Ƿ��Ǳ�ǵ�Ԫ��
    bool isMarker() const {
        return cellType != NormalCell;
    }

    // �Ƿ���Ҫ��ѯ���ձ�ģʽ��
    bool needsQuery() const {
        return cellType == DataMarker && !queryExecuted;
    }

    // �Ƿ��Ǻϲ���Ԫ�������Ԫ��
    bool isMergedMain() const {
        return mergedRange.isValid() && mergedRange.isMerged();
    }

    // ========================================
    // ===== �ı���ȡ���������ڲ�ͬ������ =====
    // ========================================

    /**
     * ��ȡ��ʾ�ı���Model��DisplayRole�ã�
     * ���ȼ�����ʽ������ > ��ʽ�ı� > ��ʾֵ
     */
    QString displayText() const {
        if (hasFormula && formulaCalculated) {
            return displayValue.toString();  // ��ʾ��ʽ������
        }
        if (hasFormula && !formulaCalculated) {
            return formula;  // ��ʾ��ʽ�ı�
        }
        return displayValue.toString();  // ��ʾ��ֵͨ��ת����ı��
    }

    /**
     * ��ȡ�༭�ı���Model��EditRole�ã�
     * ���ȼ�����ʽ > ��� > ��ʾֵ
     */
    QString editText() const {
        if (hasFormula) {
            return formula;  // �༭ʱ��ʾ��ʽ
        }
        if (!markerText.isEmpty()) {
            return markerText;  // �༭ʱ��ʾԭʼ���
        }
        // ���ݾ�����
        if (!originalMarker.isEmpty()) {
            return originalMarker;
        }
        return displayValue.toString();  // �༭��ֵͨ
    }

    /**
     * ��ȡ����ɨ����ı����������ã�
     * ���ȼ�������ı� > �����ֶ� > ��ʾֵ
     */
    QString scanText() const {
        // ����ʹ�ñ���ı�
        if (!markerText.isEmpty()) {
            return markerText;
        }
        // ���ݾ����ݣ����Դ� originalMarker ��ȡ
        if (!originalMarker.isEmpty()) {
            return originalMarker;
        }
        // ���ݾ����ݣ����Դ� value ��ȡ
        if (!value.isNull() && !value.toString().isEmpty()) {
            return value.toString();
        }
        // ��󽵼��� displayValue
        return displayValue.toString();
    }

    // ========================================
    // ===== ��ʽ�������� =====
    // ========================================

    /**
     * ���ù�ʽ
     */
    void setFormula(const QString& formulaText) {
        formula = formulaText;
        hasFormula = true;
        formulaCalculated = false;
        markerText.clear();  // ��ʽ���Ǳ��
        cellType = NormalCell;
    }
};

struct ReportColumnConfig {
    QString displayName;  // ��ʾ����
    QString rtuId;        // RTU��
    int sourceRow;        // Դ�ļ��кţ���ѡ��

    ReportColumnConfig() : sourceRow(-1) {}
};

struct HistoryReportConfig {
    QString reportName;                    // ��������
    QString configFilePath;                // �����ļ�·��
    QList<ReportColumnConfig> columns;     // �������б�
    QSet<int> dataColumns;                 // �������������ϣ����ڱ��ֻ���У�
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

// ===== ȫ�����ã��������ڿ��ܵ���չ��=====
struct GlobalDataConfig {
    TimeRangeConfig globalTimeRange;
};

#endif // DATABINDINGCONFIG_H