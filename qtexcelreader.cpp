#include <QDebug>
#include <QVariant>

#include "qtexcelreader.h"

#include "kexcelreader.h"


QtExcelReader::QtExcelReader(QWidget *parent, Qt::WindowFlags flags)
    : QMainWindow(parent, flags)
{
    ui.setupUi(this);

    connect(ui.runExcel, SIGNAL(clicked()), SLOT(slotExcelTest()));
}

QtExcelReader::~QtExcelReader()
{

}

void QtExcelReader::slotExcelTest()
{
    KExcelReader excelReader;

    QList<QVariantList> data;
    if (excelReader.open("C:\\1.xlsx"))
    {
        data = excelReader.values();
    }

    foreach(QVariantList row, data)
    {
        foreach(QVariant v, row)
        {
            bool isOk = false;
            int value = v.toInt(&isOk);
            if(isOk)
                qDebug() << value;

            isOk = false;
            QString sValue = v.toString();
            if(!isOk)
                qDebug() << sValue;
        }

//        int cn = data.count();
    }
}
