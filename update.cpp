#include "update.h"

Update::Update()
{
}

void Update::iniVersion()
{
    QString localFileName = "ServerEmployeeHttpFromSite.ini";
    fileHttpIni = new QFile;
    fileHttpIni->setFileName(localFileName);
    fileHttpIni->open(QIODevice::WriteOnly);

    url = "http://91.102.219.74/QtProject/ServerEmployee/ServerEmployee.ini";
    replyIni = httpIni.get(QNetworkRequest(url));
    connect(replyIni,SIGNAL(finished()),this,SLOT(httpDoneIni()));
    connect(replyIni,SIGNAL(readyRead()),this,SLOT(httpReadyReadIni()));
}

bool Update::newVersion()
{
    bool resultUpdates = false;
    QSettings settings("ServerEmployeeHttpFromSite.ini",QSettings::IniFormat);
    QString iniVersion = settings.value("Version").toString();

    QSettings iniSettings("ServerEmployee.ini",QSettings::IniFormat);
    nowVersion = iniSettings.value("Version").toString();

    if(iniVersion.toFloat() > nowVersion.toFloat()){
            QSettings iniSettings("ServerEmployee.ini",QSettings::IniFormat);
            iniSettings.setValue("Version",iniVersion);
            iniSettings.sync();

            resultUpdates = true;
            exeVersion();
    }
    return resultUpdates;
}

void Update::exeVersion()
{
    QString localFileName1 = "./ServerEmployee.tmp";
    fileHttpExe = new QFile(localFileName1);
    fileHttpExe->open(QIODevice::ReadWrite);

    QFile renFile;
    renFile.setFileName("./ServerEmployee.exe.bak");
    if(renFile.exists()){
        renFile.remove();
    }
    QFile fileR;
    fileR.setFileName("./ServerEmployee.exe");
    fileR.rename("./ServerEmployee.exe.bak");
    fileR.close();

//    QFile fileRe;
//    fileRe.setFileName("./ServerEmployeeHttpFromSite.ini");
//    fileRe.remove();


    url = "http://91.102.219.74/QtProject/ServerEmployee/ServerEmployee.exe";
    replyExe = httpExe.get(QNetworkRequest(url));
    connect(replyExe,SIGNAL(finished()),this,SLOT(httpDoneExe()));
    connect(replyExe,SIGNAL(readyRead()),this,SLOT(httpReadyReadExe()));
}


void Update::httpDoneIni()
{
    fileHttpIni->close();
    fileHttpIni->remove();
    newVersion();
}

void Update::httpDoneExe()
{
    fileHttpExe->rename("./ServerEmployee.exe");
    fileHttpExe->close();
}

void Update::closeConnection()
{
    httpIni.destroyed();
    httpExe.destroyed();
}

void Update::restartProgramm()
{
    QProcess::startDetached("ServerEmployee.exe");
    emit qApp->quit();
}

void Update::httpReadyReadIni()
{
    if (fileHttpIni)
        fileHttpIni->write(replyIni->readAll());
}

void Update::httpReadyReadExe()
{
    if (fileHttpExe)
        fileHttpExe->write(replyExe->readAll());
}
