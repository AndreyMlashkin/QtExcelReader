#include "qtexcelreader.h"

#include "kexcelreader.h"


QtExcelReader::QtExcelReader(QWidget *parent, Qt::WFlags flags)
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
    if (excelReader.open("D:\\1.xlsx"))
    {
        data = excelReader.values();
    }

    foreach(QVariantList row, data)
    {
        int cn = data.count();
    }
}
