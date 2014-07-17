#include <QDebug>
#include <QMessageBox>
#include <QVariant>
#include <QStandardItemModel>

#include "qtexcelreader.h"
#include "ui_qtexcelreader.h"

#include "kexcelreader.h"

const int rowCount = 10;
const int columnCount = 5;


QtExcelReader::QtExcelReader(QWidget *parent, Qt::WindowFlags flags)
    : QMainWindow(parent, flags),
      m_model(new QStandardItemModel)
{
    m_ui.setupUi(this);
    m_model->insertColumns(0, columnCount);
    m_model->insertRows(0, rowCount);
    m_ui.table->setModel(m_model);

    connect(m_ui.runExcel, SIGNAL(clicked()), SLOT(showTable()));
}

QtExcelReader::~QtExcelReader()
{

}

void QtExcelReader::showTable()
{
    KExcelReader excelReader;

    QList<QVariantList> data;
    QString path = m_ui.path->text();

    bool isOpen = (excelReader.open(path));
    if(!isOpen)
    {
        QMessageBox* b = new QMessageBox(this);
        b->setText("Документ не найден.");
        b->exec();
        return;
    }

    data = excelReader.values(columnCount, rowCount);
    int i = 0;
    foreach(QVariantList row, data)
    {
        int j = 0;
        QString str;
        foreach(QVariant v, row)
        {
            QString sValue = v.toString();
            m_model->setData(m_model->index(i, j), sValue);

            str += sValue + " ";
            ++j;
        }
        qDebug() << str;
        ++i;
    }
}
