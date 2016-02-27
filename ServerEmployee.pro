#-------------------------------------------------
#
# Project created by QtCreator 2013-12-18T09:26:54
#
#-------------------------------------------------

QT       += core gui\
            sql\
            network widgets\
            xml\
            axcontainer\

TARGET = ServerEmployee

CONFIG += console

TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \

HEADERS  += mainwindow.h \

include(Services.pri)

RESOURCES += \
    icons.qrc
