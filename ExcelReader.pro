#-------------------------------------------------
#
# Project created by QtCreator 2014-07-17T14:33:11
#
#-------------------------------------------------

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ExcelReader
TEMPLATE = app
CONFIG += qaxcontainer

SOURCES += main.cpp \
           kexcelreader.cpp \
           qtexcelreader.cpp


HEADERS += kexcelreader.h \
           qtexcelreader.h

FORMS    += qtexcelreader.ui

RESOURCES += qtexcelreader.qrc
