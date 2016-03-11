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

include(Services.pri)
include(common.pri)
include(qtservice.pri)

RESOURCES +=

RC_FILE = ServerEmployee.rc

HEADERS += \
    myservice.h

SOURCES += \
    myservice.cpp \
    main.cpp
