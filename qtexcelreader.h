#ifndef QTEXCELREADER_H
#define QTEXCELREADER_H

#include "ui_qtexcelreader.h"

class QStandardItemModel;

class QtExcelReader : public QMainWindow
{
    Q_OBJECT

public:
    QtExcelReader(QWidget *parent = 0, Qt::WindowFlags flags = 0);
    ~QtExcelReader();

protected slots:
    void showTable();

private:
    Ui::QtExcelReaderClass m_ui;
    QStandardItemModel* m_model;
};

#endif // QTEXCELREADER_H
