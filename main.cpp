#include "qtexcelreader.h"
//#include <QtGui/QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    QtExcelReader w;
    w.show();
    return a.exec();
}
