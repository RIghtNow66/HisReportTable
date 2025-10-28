#include <QApplication>
#include "ReportTableWidget.h"

int main(int argc, char* argv[]){
    QApplication app(argc, argv);

    ReportTableWidget w;
    w.show();

    return app.exec();
}

