TEMPLATE = app
LANGUAGE = C++
TARGET = ScadaReportControl

# ============ 项目配置 ============
include($(DEVHOME)/source/include/projectdef.pro)

QMAKE_CXXFLAGS += /D_HAS_VECTOR_ALGORITHMS=0

# 覆盖 debug 设置（允许 debug 和 release）
CONFIG -= debug release
CONFIG += debug_and_release

# Qt 模块
QT += core widgets gui core-private gui-private svg concurrent printsupport

# ============ 路径设置 ============
INCLUDEPATH += $${APP_INC}
INCLUDEPATH += $$PWD/include
DEPENDPATH += $$PWD/include

LIBS += -liosal -ltaos -litaosdbms

# 库文件搜索路径
LIBS += -L$$PWD/lib

# ============ 链接库 ============

# 使用更可靠的判断方式
contains(QMAKE_TARGET.arch, x86_64) {
    CONFIG(release, debug|release) {
        LIBS += -lQXlsx
        message("Linking QXlsx (Release)")
    } else {
        LIBS += -lQXlsxd
        message("Linking QXlsxd (Debug)")
    }
}

# ============ 源文件 ============
SOURCES += \
    main.cpp \
    mainwindow.cpp \
    reportdatamodel.cpp \
    formulaengine.cpp \
    excelhandler.cpp \
    EnhancedTableView.cpp \
    TaosDataFetcher.cpp \
    DayReportParser.cpp \
    BaseReportParser.cpp \
    MonthReportParser.cpp\
	TimeSettingsDialog.cpp\
	UnifiedQueryParser.cpp\

# ============ 头文件 ============
HEADERS += \
    mainwindow.h \
    reportdatamodel.h \
    formulaengine.h \
    excelhandler.h \
    DataBindingConfig.h \
    EnhancedTableView.h \
    TaosDataFetcher.h \
    DayReportParser.h \
    BaseReportParser.h \
    MonthReportParser.h\
	TimeSettingsDialog.h\
	UnifiedQueryParser.h\

# ============ 资源文件 ============
RESOURCES += ReportTable.qrc