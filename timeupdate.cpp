#include "timeupdate.h"
//#include "exportxml.h"

TimeUpdate::TimeUpdate()
{
    exportXML.openImport(true);

    QTimer *timer = new QTimer(this);
    connect(timer,SIGNAL(timeout()),this,SLOT(updateTime()));
    timer->start(5400000);
}

TimeUpdate::~TimeUpdate()
{
}

void TimeUpdate::updateTime()
{
    //QTime currTime = QTime::currentTime();

    //if((currTime > QTime(12,00) && currTime < QTime(13,00)) || (currTime > QTime(21,00) && currTime < QTime(22,00))){
    //    qDebug()<<"All Data start";

    //    exportXML.openImport(true);
    //}else{
        updateFtp.iniVersion();
        qDebug()<<"All Data...";

        QMutex mutex;
        mutex.lock();

        QWaitCondition pause;
        pause.wait(&mutex, 60000);

        exportXML.openImport(true);
    //}
}
