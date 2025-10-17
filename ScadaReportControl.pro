TEMPLATE = app
LANGUAGE = C++
TARGET = ScadaReportControl

include(QXlsx/QXlsx/QXlsx.pri)
include( $(DEVHOME)/source/include/projectdef.pro )
LIBS += -liosal  -ltaos -litaosdbms
INCLUDEPATH += $${APP_INC}

QT += core widgets gui core-private gui-private svg concurrent

SOURCES += \
    main.cpp\
    mainwindow.cpp\
    reportdatamodel.cpp\
    formulaengine.cpp\
    excelhandler.cpp\
	EnhancedTableView.cpp\
	TaosDataFetcher.cpp\
	DayReportParser.cpp\
	BaseReportParser.cpp\
	MonthReportParser.cpp\
	
	

HEADERS +=\
    mainwindow.h\
    reportdatamodel.h\
    formulaengine.h\
    excelhandler.h\
    DataBindingConfig.h\
	EnhancedTableView.h\
	TaosDataFetcher.h\
	DayReportParser.h\
	BaseReportParser.h\
	MonthReportParser.h\


RESOURCES += ReportTable.qrc
