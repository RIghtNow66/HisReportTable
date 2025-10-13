TEMPLATE = app
LANGUAGE = C++
TARGET = ScadaReportControl

include(QXlsx/QXlsx/QXlsx.pri)
include( $(DEVHOME)/source/include/projectdef.pro )
LIBS += -liosal -ligdbi -lihmiapi -linetapi -lirtdbapi  -ltaos -litaosdbms
INCLUDEPATH += $${APP_INC}

QT += core widgets gui core-private gui-private svg

SOURCES += \
    main.cpp\
    mainwindow.cpp\
    reportdatamodel.cpp\
    formulaengine.cpp\
    excelhandler.cpp\
	EnhancedTableView.cpp\
	UniversalQueryEngine.cpp\
	TaosDataFetcher.cpp\
	DayReportParser.cpp\
	

HEADERS +=\
    mainwindow.h\
    reportdatamodel.h\
    formulaengine.h\
    excelhandler.h\
    DataBindingConfig.h\
	EnhancedTableView.h\
	UniversalQueryEngine.h\
	TaosDataFetcher.h\
	DayReportParser.h\


RESOURCES += ReportTable.qrc
