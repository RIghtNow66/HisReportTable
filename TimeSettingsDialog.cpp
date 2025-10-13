// TimeSettingsDialog.cpp
#include "TimeSettingsDialog.h"
#include <QGridLayout>
#include <QFrame>

TimeSettingsDialog::TimeSettingsDialog(QWidget* parent)
    : QDialog(parent)
    , m_currentType(Daily)
    , m_intervalValue(5)
    , m_intervalUnit(1) // 默认分钟
    , m_singlePointMode(false)  // 新增：单点查询模式标志
{
    setWindowTitle("时间设置");
    setModal(true);
    setMinimumWidth(500);

    // 初始化默认时间：今日00:00
    QDateTime now = QDateTime::currentDateTime();
    m_startTime = QDateTime(now.date(), QTime(0, 0, 0));

    setupUI();

    // 初始化计算终止时间
    calculateEndTime();
    adjustIntervalForReportType(); // 初始化时调整间隔
    updateIntervalDisplay();
}

void TimeSettingsDialog::setupUI()
{
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(15);
    mainLayout->setContentsMargins(20, 20, 20, 20);

    // ===== 报表类型选择 =====
    QGroupBox* typeGroup = new QGroupBox("报表类型");
    QHBoxLayout* typeLayout = new QHBoxLayout(typeGroup);

    m_reportTypeGroup = new QButtonGroup(this);

    m_dailyRadio = new QRadioButton("日报");
    m_weeklyRadio = new QRadioButton("周报");
    m_monthlyRadio = new QRadioButton("月报");
    m_customRadio = new QRadioButton("自定义");
    m_singlePointRadio = new QRadioButton("单点查询");  // 新增

    m_dailyRadio->setChecked(true);

    m_reportTypeGroup->addButton(m_dailyRadio, Daily);
    m_reportTypeGroup->addButton(m_weeklyRadio, Weekly);
    m_reportTypeGroup->addButton(m_monthlyRadio, Monthly);
    m_reportTypeGroup->addButton(m_customRadio, Custom);
    m_reportTypeGroup->addButton(m_singlePointRadio, SinglePoint);  // 新增

    typeLayout->addWidget(m_dailyRadio);
    typeLayout->addWidget(m_weeklyRadio);
    typeLayout->addWidget(m_monthlyRadio);
    typeLayout->addWidget(m_customRadio);
    typeLayout->addWidget(m_singlePointRadio);  // 新增
    typeLayout->addStretch();

    mainLayout->addWidget(typeGroup);

    connect(m_reportTypeGroup, QOverload<int>::of(&QButtonGroup::idClicked),
        this, &TimeSettingsDialog::onReportTypeChanged);

    // ===== 时间范围 =====
    QGroupBox* timeGroup = new QGroupBox("时间范围");
    QGridLayout* timeLayout = new QGridLayout(timeGroup);
    timeLayout->setColumnStretch(1, 1);

    // 起始时间
    QLabel* startLabel = new QLabel("查询时间：");  // 修改标签
    m_startTimeEdit = new QDateTimeEdit();
    m_startTimeEdit->setDisplayFormat("yyyy-MM-dd HH:mm:ss");
    m_startTimeEdit->setCalendarPopup(true);
    m_startTimeEdit->setDateTime(m_startTime);

    timeLayout->addWidget(startLabel, 0, 0);
    timeLayout->addWidget(m_startTimeEdit, 0, 1);

    connect(m_startTimeEdit, &QDateTimeEdit::dateTimeChanged,
        this, &TimeSettingsDialog::onStartTimeChanged);

    // 终止时间
    m_endLabel = new QLabel("终止时间：");  // 保存标签引用
    m_endTimeEdit = new QDateTimeEdit();
    m_endTimeEdit->setDisplayFormat("yyyy-MM-dd HH:mm:ss");
    m_endTimeEdit->setCalendarPopup(true);
    m_endTimeEdit->setReadOnly(true);
    m_endTimeEdit->setButtonSymbols(QAbstractSpinBox::NoButtons);
    m_endTimeEdit->setStyleSheet("QDateTimeEdit { background-color: #f0f0f0; }");

    timeLayout->addWidget(m_endLabel, 1, 0);
    timeLayout->addWidget(m_endTimeEdit, 1, 1);

    connect(m_endTimeEdit, &QDateTimeEdit::dateTimeChanged,
        this, &TimeSettingsDialog::onEndTimeChanged);

    mainLayout->addWidget(timeGroup);

    // ===== 间隔时间 =====
    m_intervalGroup = new QGroupBox("时间间隔");  // 保存引用
    QVBoxLayout* intervalMainLayout = new QVBoxLayout(m_intervalGroup);

    QHBoxLayout* intervalInputLayout = new QHBoxLayout();

    QLabel* intervalLabel = new QLabel("间隔：");
    m_intervalSpinBox = new QSpinBox();
    m_intervalSpinBox->setRange(1, 99999);
    m_intervalSpinBox->setValue(m_intervalValue);
    m_intervalSpinBox->setMinimumWidth(80);

    m_intervalUnitCombo = new QComboBox();
    m_intervalUnitCombo->addItem("秒");
    m_intervalUnitCombo->addItem("分钟");
    m_intervalUnitCombo->addItem("小时");
    m_intervalUnitCombo->addItem("天");
    m_intervalUnitCombo->setCurrentIndex(m_intervalUnit);

    m_intervalDisplayLabel = new QLabel();
    m_intervalDisplayLabel->setStyleSheet("QLabel { color: #666; font-size: 11px; }");

    intervalInputLayout->addWidget(intervalLabel);
    intervalInputLayout->addWidget(m_intervalSpinBox);
    intervalInputLayout->addWidget(m_intervalUnitCombo);
    intervalInputLayout->addWidget(m_intervalDisplayLabel);
    intervalInputLayout->addStretch();

    intervalMainLayout->addLayout(intervalInputLayout);

    connect(m_intervalSpinBox, QOverload<int>::of(&QSpinBox::valueChanged),
        this, &TimeSettingsDialog::onIntervalValueChanged);
    connect(m_intervalUnitCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
        this, &TimeSettingsDialog::onIntervalUnitChanged);

    // 快捷按钮
    QHBoxLayout* quickLayout = new QHBoxLayout();
    QLabel* quickLabel = new QLabel("常用间隔：");

    m_quick5SecBtn = new QPushButton("5秒");
    m_quick1MinBtn = new QPushButton("1分钟");
    m_quick5MinBtn = new QPushButton("5分钟");
    m_quick1HourBtn = new QPushButton("1小时");

    m_quick5SecBtn->setMaximumWidth(70);
    m_quick1MinBtn->setMaximumWidth(70);
    m_quick5MinBtn->setMaximumWidth(70);
    m_quick1HourBtn->setMaximumWidth(70);

    connect(m_quick5SecBtn, &QPushButton::clicked, this, &TimeSettingsDialog::onQuickIntervalClicked);
    connect(m_quick1MinBtn, &QPushButton::clicked, this, &TimeSettingsDialog::onQuickIntervalClicked);
    connect(m_quick5MinBtn, &QPushButton::clicked, this, &TimeSettingsDialog::onQuickIntervalClicked);
    connect(m_quick1HourBtn, &QPushButton::clicked, this, &TimeSettingsDialog::onQuickIntervalClicked);

    quickLayout->addWidget(quickLabel);
    quickLayout->addWidget(m_quick5SecBtn);
    quickLayout->addWidget(m_quick1MinBtn);
    quickLayout->addWidget(m_quick5MinBtn);
    quickLayout->addWidget(m_quick1HourBtn);
    quickLayout->addStretch();

    intervalMainLayout->addLayout(quickLayout);
    mainLayout->addWidget(m_intervalGroup);

    // ===== 底部按钮 =====
    mainLayout->addStretch();

    QHBoxLayout* buttonLayout = new QHBoxLayout();
    buttonLayout->addStretch();

    m_okBtn = new QPushButton("确定");
    m_cancelBtn = new QPushButton("取消");

    m_okBtn->setMinimumWidth(80);
    m_cancelBtn->setMinimumWidth(80);

    buttonLayout->addWidget(m_okBtn);
    buttonLayout->addWidget(m_cancelBtn);

    mainLayout->addLayout(buttonLayout);

    connect(m_okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(m_cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
}

void TimeSettingsDialog::onReportTypeChanged(int id)
{
    m_currentType = static_cast<ReportType>(id);

    // 更新单点查询模式标志
    m_singlePointMode = (m_currentType == SinglePoint);

    // 根据模式显示/隐藏控件
    updateUIForMode();

    if (m_currentType != Custom && m_currentType != SinglePoint) {
        calculateEndTime();
    }

    if (m_currentType != SinglePoint) {
        adjustIntervalForReportType();
    }
}

void TimeSettingsDialog::updateUIForMode()
{
    if (m_singlePointMode) {
        // 单点模式：隐藏终止时间和间隔设置
        m_endLabel->setVisible(false);
        m_endTimeEdit->setVisible(false);
        m_intervalGroup->setVisible(false);
    }
    else {
        // 其他模式：显示所有控件
        m_endLabel->setVisible(true);
        m_endTimeEdit->setVisible(true);
        m_intervalGroup->setVisible(true);
        updateEndTimeEditability();
    }
}

void TimeSettingsDialog::adjustIntervalForReportType()
{
    switch (m_currentType) {
    case Daily:
        m_intervalSpinBox->setValue(5);
        m_intervalUnitCombo->setCurrentIndex(1);
        m_quick5SecBtn->setText("10秒");
        m_quick1MinBtn->setText("1分钟");
        m_quick5MinBtn->setText("5分钟");
        m_quick1HourBtn->setText("30分钟");
        break;

    case Weekly:
        m_intervalSpinBox->setValue(1);
        m_intervalUnitCombo->setCurrentIndex(2);
        m_quick5SecBtn->setText("30分钟");
        m_quick1MinBtn->setText("1小时");
        m_quick5MinBtn->setText("3小时");
        m_quick1HourBtn->setText("6小时");
        break;

    case Monthly:
        m_intervalSpinBox->setValue(1);
        m_intervalUnitCombo->setCurrentIndex(3);
        m_quick5SecBtn->setText("6小时");
        m_quick1MinBtn->setText("12小时");
        m_quick5MinBtn->setText("1天");
        m_quick1HourBtn->setText("2天");
        break;

    case Custom:
        m_quick5SecBtn->setText("5秒");
        m_quick1MinBtn->setText("1分钟");
        m_quick5MinBtn->setText("5分钟");
        m_quick1HourBtn->setText("1小时");
        break;

    case SinglePoint:
        // 单点模式不需要设置间隔
        break;
    }

    if (m_currentType != SinglePoint) {
        updateIntervalDisplay();
    }
}



void TimeSettingsDialog::onStartTimeChanged(const QDateTime& dateTime)
{
    m_startTime = dateTime;

    if (m_currentType == SinglePoint) {
        // 单点模式：直接返回，不需要计算终止时间
        return;
    }

    // 非自定义模式下自动计算终止时间
    if (m_currentType != Custom) {
        calculateEndTime();
    }
    // 自定义模式下，如果终止时间早于起始时间，自动调整
    else if (m_endTime < m_startTime) {
        m_endTime = m_startTime.addSecs(3600); // 默认加1小时
        m_endTime = limitToCurrentTime(m_endTime);
        m_endTimeEdit->setDateTime(m_endTime);
    }
}

void TimeSettingsDialog::onEndTimeChanged(const QDateTime& dateTime)
{
    // 只在自定义模式下响应手动修改
    if (m_currentType == Custom) {
        m_endTime = dateTime;

        // 确保终止时间不早于起始时间
        if (m_endTime < m_startTime) {
            m_endTime = m_startTime;
            m_endTimeEdit->setDateTime(m_endTime);
        }

        // 确保不超过当前时间
        m_endTime = limitToCurrentTime(m_endTime);
        if (m_endTime != dateTime) {
            m_endTimeEdit->setDateTime(m_endTime);
        }
    }
}

void TimeSettingsDialog::updateEndTimeEditability()
{
    if (m_currentType == Custom) {
        // 自定义模式：可编辑
        m_endTimeEdit->setReadOnly(false);
        m_endTimeEdit->setButtonSymbols(QAbstractSpinBox::UpDownArrows);
        m_endTimeEdit->setStyleSheet("QDateTimeEdit { background-color: white; }");
    }
    else {
        // 其他模式：只读
        m_endTimeEdit->setReadOnly(true);
        m_endTimeEdit->setButtonSymbols(QAbstractSpinBox::NoButtons);
        m_endTimeEdit->setStyleSheet("QDateTimeEdit { background-color: #f0f0f0; }");
    }
}

void TimeSettingsDialog::calculateEndTime()
{
    QDateTime calculatedEnd;

    switch (m_currentType) {
    case Daily:
        calculatedEnd = m_startTime.addDays(1);
        break;

    case Weekly:
        calculatedEnd = m_startTime.addDays(7);
        break;

    case Monthly:
        calculatedEnd = m_startTime.addDays(30);
        break;

    case Custom:
    case SinglePoint:
        return;
    }

    m_endTime = limitToCurrentTime(calculatedEnd);
    m_endTimeEdit->setDateTime(m_endTime);
}

QDateTime TimeSettingsDialog::limitToCurrentTime(const QDateTime& time)
{
    QDateTime now = QDateTime::currentDateTime();
    return (time > now) ? now : time;
}

void TimeSettingsDialog::onIntervalValueChanged(int value)
{
    m_intervalValue = value;
    updateIntervalDisplay();
}

void TimeSettingsDialog::onIntervalUnitChanged(int index)
{
    m_intervalUnit = index;
    updateIntervalDisplay();
}

void TimeSettingsDialog::updateIntervalDisplay()
{
    int totalSeconds = getIntervalSeconds();

    QString displayText;
    if (totalSeconds < 60) {
        displayText = QString("（%1秒）").arg(totalSeconds);
    }
    else if (totalSeconds < 3600) {
        displayText = QString("（%1分钟）").arg(totalSeconds / 60);
    }
    else if (totalSeconds < 86400) {
        displayText = QString("（%1小时）").arg(totalSeconds / 3600);
    }
    else {
        displayText = QString("（%1天）").arg(totalSeconds / 86400);
    }

    m_intervalDisplayLabel->setText(displayText);
}

void TimeSettingsDialog::onQuickIntervalClicked()
{
    QPushButton* btn = qobject_cast<QPushButton*>(sender());
    if (!btn) return;

    QString btnText = btn->text();

    if (btnText == "5秒" || btnText == "10秒") {
        m_intervalSpinBox->setValue(btnText == "5秒" ? 5 : 10);
        m_intervalUnitCombo->setCurrentIndex(0);
    }
    else if (btnText == "1分钟" || btnText == "30分钟") {
        m_intervalSpinBox->setValue(btnText == "1分钟" ? 1 : 30);
        m_intervalUnitCombo->setCurrentIndex(1);
    }
    else if (btnText == "5分钟" || btnText == "3小时" || btnText == "12小时") {
        if (btnText == "5分钟") {
            m_intervalSpinBox->setValue(5);
            m_intervalUnitCombo->setCurrentIndex(1);
        }
        else if (btnText == "3小时") {
            m_intervalSpinBox->setValue(3);
            m_intervalUnitCombo->setCurrentIndex(2);
        }
        else {
            m_intervalSpinBox->setValue(12);
            m_intervalUnitCombo->setCurrentIndex(2);
        }
    }
    else if (btnText == "1小时" || btnText == "6小时") {
        m_intervalSpinBox->setValue(btnText == "1小时" ? 1 : 6);
        m_intervalUnitCombo->setCurrentIndex(2);
    }
    else if (btnText == "1天" || btnText == "2天") {
        m_intervalSpinBox->setValue(btnText == "1天" ? 1 : 2);
        m_intervalUnitCombo->setCurrentIndex(3);
    }
}

// ===== 公共接口实现 =====

QDateTime TimeSettingsDialog::getStartTime() const
{
    return m_startTime;
}

QDateTime TimeSettingsDialog::getEndTime() const
{
    if (m_singlePointMode) {
        return m_startTime;  // 单点模式返回相同的时间
    }
    return m_endTime;
}

int TimeSettingsDialog::getIntervalSeconds() const
{
    if (m_singlePointMode) {
        return 0;  // 单点模式间隔为0
    }

    int multiplier = 1;
    switch (m_intervalUnit) {
    case 0: multiplier = 1; break;
    case 1: multiplier = 60; break;
    case 2: multiplier = 3600; break;
    case 3: multiplier = 86400; break;
    }
    return m_intervalValue * multiplier;
}


TimeSettingsDialog::ReportType TimeSettingsDialog::getReportType() const
{
    return m_currentType;
}

bool TimeSettingsDialog::isSinglePointMode() const
{
    return m_singlePointMode;
}

void TimeSettingsDialog::setStartTime(const QDateTime& time)
{
    m_startTime = time;
    m_startTimeEdit->setDateTime(time);
    if (!m_singlePointMode) {
        calculateEndTime();
    }
}

void TimeSettingsDialog::setReportType(ReportType type)
{
    m_currentType = type;
    m_singlePointMode = (type == SinglePoint);

    switch (type) {
    case Daily:
        m_dailyRadio->setChecked(true);
        break;
    case Weekly:
        m_weeklyRadio->setChecked(true);
        break;
    case Monthly:
        m_monthlyRadio->setChecked(true);
        break;
    case Custom:
        m_customRadio->setChecked(true);
        break;
    case SinglePoint:
        m_singlePointRadio->setChecked(true);
        break;
    }

    updateUIForMode();
    if (!m_singlePointMode) {
        updateEndTimeEditability();
        calculateEndTime();
    }
}