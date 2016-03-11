#include "myservice.h"
#include "update.h"
#include "TimeUpdate.h"

MyService::MyService(int argc, char **argv) : QtService<QCoreApplication>(argc,argv,"MyService")
{
    try {
        setServiceDescription("Server Employee 1C");
        setServiceFlags(QtServiceBase::CanBeSuspended);
//        Update upDate;
//        upDate.iniVersion();

        TimeUpdate timeUpdate;
    } catch (...) {
        qCritical() << "An unknown error in constructor";
    }
}

MyService::~MyService()
{
    try {

    } catch (...) {
        qCritical() << "An unknown error in destructor";
    }
}

void MyService::start()
{
    try {
        QCoreApplication *app = application();
        qDebug() << "Service started!";
        qDebug() << app->applicationDirPath();
        //My Class
    } catch (...) {
        qCritical() << "An unknown start error";
    }
}

void MyService::pause()
{
    try {
        qDebug() << "Service paused!";
    } catch (...) {
        qCritical() << "An unknown pause error";
    }
}

void MyService::resume()
{
    try {
        qDebug() << "Service resume!";
    } catch (...) {
        qCritical() << "An unknown resume error";
    }
}

void MyService::stop()
{
    try {
        qDebug() << "Service stopped!";
    } catch (...) {
        qCritical() << "An unknown stop error";
    }
}

