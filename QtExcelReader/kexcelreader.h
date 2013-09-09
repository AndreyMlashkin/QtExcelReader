#ifndef KEXCELREADER_H
#define KEXCELREADER_H

#include <QObject>
#include <QAxObject>

class KExcelReader : public QObject
{
    Q_OBJECT

public:
    KExcelReader(QObject *parent=0);
    ~KExcelReader();

public:
    bool open(const QString& xlsFile);
    int sheetCount();
    int rowCount(QAxObject* sheet);
    int columnCount(QAxObject* sheet);
    QList<QVariantList> values(int sheetNumber=1);

private:
    void init();

private:
    QAxObject *m_excel;
    QAxObject *m_workbooks;
    QAxObject *m_workbook;
    QAxObject *m_sheets;
};

#endif // KEXCELREADER_H
