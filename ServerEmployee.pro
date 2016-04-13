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
            #axserver\

TARGET = ServerEmployee

CONFIG += console

TEMPLATE = app

include(Services.pri)

RESOURCES +=

RC_FILE = ServerEmployee.rc

HEADERS +=

SOURCES += \
    main.cpp
