#ifndef QTEXCELREADER_H
#define QTEXCELREADER_H

#include <QtGui/QMainWindow>
#include "ui_qtexcelreader.h"

class QtExcelReader : public QMainWindow
{
    Q_OBJECT

public:
    QtExcelReader(QWidget *parent = 0, Qt::WFlags flags = 0);
    ~QtExcelReader();

protected slots:
    void slotExcelTest();

private:
    Ui::QtExcelReaderClass ui;
};

#endif // QTEXCELREADER_H
