#include "update.h"

Update::Update(QWidget *parent) :
    QDialog(parent)
{
    QFile file(":/QWidgetStyle.txt");
    file.open(QFile::ReadOnly);
    QString styleSheetString = QLatin1String(file.readAll());
    QWidget::setStyleSheet(styleSheetString);
}

void Update::iniVersion()
{
    QString localFileName = "ServerEmployeeHttpFromSite.ini";
    if(localFileName.isEmpty())
        localFileName = "null.out";
    fileHttpIni = new QFile;
    fileHttpIni->setFileName(localFileName);
    if(!fileHttpIni->open(QIODevice::WriteOnly)){
        QMessageBox::warning(this,tr("Attention"),tr("Don't open %1").arg(localFileName));
    }

    url = "http://91.102.219.74/QtProject/ServerEmployee/ServerEmployee.ini";//url.toString() += "EmployeeClient.ini";
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
    progressDialogExe = new QProgressDialog(this);
    connect(progressDialogExe,SIGNAL(canceled()),this,SLOT(cancelDownLoadExe()));
    QString localFileName1 = "./ServerEmployee.tmp";
    if(localFileName1.isEmpty())
        localFileName1 = "null.out";
    fileHttpExe = new QFile(localFileName1);
    if(!fileHttpExe->open(QIODevice::ReadWrite)){

    }

    QFile renFile;
    renFile.setFileName("./ServerEmployee.exe.bak");
    if(renFile.exists()){
        renFile.remove();
    }
    QFile fileR;
    fileR.setFileName("./ServerEmployee.exe");
    fileR.rename("./ServerEmployee.exe.bak");
    fileR.close();

    QFile fileRe;
    fileRe.setFileName("./ServerEmployeeHttpFromSite.ini");
    fileRe.remove();


    url = "http://91.102.219.74/QtProject/ServerEmployee/ServerEmployee.exe";//url.toString() += "EmployeeClient.exe";
    replyExe = httpExe.get(QNetworkRequest(url));
    connect(replyExe,SIGNAL(finished()),this,SLOT(httpDoneExe()));
    connect(replyExe,SIGNAL(readyRead()),this,SLOT(httpReadyReadExe()));
    connect(replyExe,SIGNAL(uploadProgress(qint64,qint64)),this,SLOT(updateDataReadProgressExe(qint64,qint64)));

    progressDialogExe->setLabelText(tr("Downloading %1 ...").arg(localFileName1));
    progressDialogExe->setEnabled(true);
    progressDialogExe->exec();
}

void Update::tranceVersion()
{
    progressDialogTrance = new QProgressDialog(this);
    connect(progressDialogTrance,SIGNAL(canceled()),this,SLOT(cancelDownLoadTrance()));
    QString localFileName1 = "./ServerEmployee_ru.qm";
    if(localFileName1.isEmpty())
        localFileName1 = "null.out";
    fileHttpTrance = new QFile(localFileName1);
    if(!fileHttpTrance->open(QIODevice::WriteOnly)){

    }
    url = "http://91.102.219.74/QtProject/ServerEmployee/ServerEmployee_ru.qm";//url.toString() += "EmployeeClient_ru.qm";

    replyTrance = httpTrance.get(QNetworkRequest(url));
    connect(replyTrance,SIGNAL(finished()),this,SLOT(httpDoneTrance()));
    connect(replyTrance,SIGNAL(readyRead()),this,SLOT(httpReadyReadTrance()));
    connect(replyTrance,SIGNAL(uploadProgress(qint64,qint64)),this,SLOT(updateDataReadProgressTrance(qint64,qint64)));

    progressDialogTrance->setLabelText(tr("Downloading %1 ...").arg(localFileName1));
    progressDialogTrance->setEnabled(false);
    progressDialogTrance->exec();
}

void Update::httpDoneIni()
{
    fileHttpIni->close();
    newVersion();
}

void Update::httpDoneExe()
{
    fileHttpExe->rename("./ServerEmployee.exe");
    fileHttpExe->close();
    tranceVersion();
}

void Update::httpDoneTrance()
{
    fileHttpTrance->close();
    closeConnection();
    restartProgramm();
}

void Update::closeConnection()
{
    httpIni.destroyed();
    httpExe.destroyed();
    httpTrance.destroyed();
}

void Update::cancelDownLoadIni()
{
    if(fileHttpIni->exists()){
        fileHttpIni->close();
        fileHttpIni->remove();
    }
}

void Update::cancelDownLoadExe()
{
    if(fileHttpExe->exists()){

        fileHttpExe->close();
        fileHttpExe->remove();
    }
}

void Update::cancelDownLoadTrance()
{
    if(fileHttpTrance->exists()){
        fileHttpTrance->close();
        fileHttpTrance->remove();
    }
}

void Update::restartProgramm()
{
    QProcess::startDetached("ServerEmployee.exe");
    QApplication::setQuitOnLastWindowClosed(true);
    emit qApp->quit();
}


void Update::updateDataReadProgressExe(qint64 readBytes, qint64 totalBytes)
{
    progressDialogExe->setMaximum(totalBytes);
    progressDialogExe->setValue(readBytes);
}

void Update::updateDataReadProgressTrance(qint64 readBytes, qint64 totalBytes)
{
    progressDialogTrance->setMaximum(totalBytes);
    progressDialogTrance->setValue(readBytes);
}

void Update::httpReadyReadIni()
{
    if (fileHttpIni)
        fileHttpIni->write(replyIni->readAll());
}

void Update::httpReadyReadExe()
{
    if (fileHttpExe){
        fileHttpExe->write(replyExe->readAll());

    }
    progressDialogExe->hide();
}

void Update::httpReadyReadTrance()
{
    if (fileHttpTrance)
        fileHttpTrance->write(replyTrance->readAll());
    progressDialogTrance->hide();
}
