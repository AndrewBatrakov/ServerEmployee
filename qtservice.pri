include(../common.pri)
INCLUDEPATH += $$PWD
DEPENDPATH += $$PWD
!win32:QT += network
win32:LIBS += -luser32

qtservice-uselib:!qtservice-buildlib {
    LIBS += -L$$QTSERVICE_LIBDIR -l$$QTSERVICE_LIBNAME
} else {
    HEADERS       +=
    SOURCES       +=
    win32:SOURCES +=
    unix:HEADERS  +=
    unix:SOURCES  +=
}

win32 {
    qtservice-buildlib:shared:DEFINES += QT_QTSERVICE_EXPORT
    else:qtservice-uselib:DEFINES += QT_QTSERVICE_IMPORT
}
