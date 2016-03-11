#include "timeupdate.h"
#include "exportxml.h"

TimeUpdate::TimeUpdate()
{
    QTimer *timer = new QTimer(this);
    connect(timer,SIGNAL(timeout()),this,SLOT(updateTime()));
    timer->start(1000);
}

TimeUpdate::~TimeUpdate()
{
}

void TimeUpdate::updateTime()
{
    QString currTime = QTime::currentTime().toString();
    if(currTime == "00:00:00"){
        ExportXML exportXML;
        exportXML.openImport(true);
    }else if(currTime == "03:00:00"){
        ExportXML exportXML;
        exportXML.openImport(false);
    }else if(currTime == "06:00:00"){
        ExportXML exportXML;
        exportXML.openImport(false);
    }else if(currTime == "09:00:00"){
        ExportXML exportXML;
        exportXML.openImport(false);
    }else if(currTime == "12:00:00"){
        ExportXML exportXML;
        exportXML.openImport(true);
    }else if(currTime == "15:00:00"){
        ExportXML exportXML;
        exportXML.openImport(false);
    }else if(currTime == "18:00:00"){
        ExportXML exportXML;
        exportXML.openImport(false);
    }else if(currTime == "21:00:00"){
        ExportXML exportXML;
        exportXML.openImport(true);
    }

    if(currTime == "23:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "02:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "05:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "08:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "11:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "14:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "17:30:00"){
        updateFtp.iniVersion();
    }else if(currTime == "20:30:00"){
        updateFtp.iniVersion();
    }
}
