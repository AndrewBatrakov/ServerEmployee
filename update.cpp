#include "update.h"
//#include <QApplication>
//#include "mainwindow.h"

Update::Update(QWidget *parent) :
    QDialog(parent)
{
   QWidget::setStyleSheet("MainWindow, QMessageBox, QDialog, QMenu, QAction, QMenuBar {background-color: "
                           "#DDD6FF}"
                           "QMenu {"
                           "font: bold}"
                           "QMenu::item:selected {background-color: "
                           "qlineargradient(x1:0, y1:0, x2:0, y2:1,"
                           "stop:0 #cfcccc,"
                           "stop:0.5 #333232,"
                           "stop:0.51 #000000,"
                           "stop:1 #585c5f);"
                           "color: #00cc00}"

                           "QMenuBar::item:selected {background-color: "
                           "qlineargradient(x1:0, y1:0, x2:0, y2:1,"
                           "stop:0 #cfcccc,"
                           "stop:0.5 #333232,"
                           "stop:0.51 #000000,"
                           "stop:1 #585c5f);"
                           "color: #00cc00}"

                           "QPushButton {"
                           "border: 1px solid black;"
                           "min-height: 20px;"
                           "min-width: 70px;"
                           "padding: 1px;"
                           "border-top-right-radius: 5px;"
                           "border-top-left-radius: 5px;"
                           "border-bottom-right-radius: 5px;"
                           "border-bottom-left-radius: 5px;"
                           "background-color: "
                           "qlineargradient(x1:0, y1:0, x2:0, y2:1,"
                           "stop:0 #d3d3d3,"
                           "stop:0.5 #bebebe,"
                           "stop:0.51 #bebebe,"
                           "stop:1 #848484);"
                           "color: #231A4C;"
                           "font: bold}"

                           "QPushButton:hover {background-color: "
                           "qlineargradient(x1:0, y1:0, x2:0, y2:1,"
                           "stop:0 #cfcccc,"
                           "stop:0.5 #333232,"
                           "stop:0.51 #000000,"
                           "stop:1 #585c5f);"
                           "color: #00cc00}"

                           "QPushButton:focus {border-color: yellow}"
                           "LineEdit:hover {background-color: #FFFF99}"
                           "QComboBox:hover {background-color: #FFFF99}"
                           "QDateEdit:hover {background-color: #FFFF99}"
                           "LineEdit:focus {background-color: #FFFFCC}"
                           "QComboBox:focus {background-color: #FFFFCC}"
                           "QDateEdit:focus {background-color: #FFFFCC}"

                           "QProgressBar {"
                           "border: 1px solid black;"
                           "text-align: center;"
                           "color: #00B600;"
                           "font: bold;"
                           "padding: 1px;"
                           "border-top-right-radius: 5px;"
                           "border-top-left-radius: 5px;"
                           "border-bottom-right-radius: 5px;"
                           "border-bottom-left-radius: 5px;"
                           "background: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1,"
                           "stop: 0 #fff,"
                           "stop: 0.4999 #eee,"
                           "stop: 0.5 #ddd,"
                           "stop: 1 #eee );"
                           "width: 15px;"
                           "}"

                           "QProgressBar::chunk {"
                           "background: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1,"
                           "stop: 0 #78d,"
                           "stop: 0.4999 #46a,"
                           "stop: 0.5 #45a,"
                           "stop: 1 #238 );"
                           "border-top-right-radius: 5px;"
                           "border-top-left-radius: 5px;"
                           "border-bottom-right-radius: 5px;"
                           "border-bottom-left-radius: 5px;"
                           "border: 1px solid black;}"

                           "QTabWidget::pane {"
                           "border: 1px solid #A3A3FF;"
                           "border-top-right-radius: 5px;"
                           "border-top-left-radius: 5px;"
                           "border-bottom-right-radius: 5px;"
                           "border-bottom-left-radius: 5px;"
                           "top: -0.5em}"

                           "QTabWidget::tab-bar {"
                           "left: 5px;}"

                           "QTabBar::tab {background-color: "
                           "qlineargradient(x1:0, y1:0, x2:0, y2:1,"
                           "stop:0 #DDD6FF,"
                           "stop:0.5 #A3A3FF,"
                           "stop:0.51 #A3A3FF,"
                           "stop:1 #DDD6FF);"
                           "border: 1px solid #A3A3FF;"
                           "padding: 1px;"
                           "min-width: 90px;"
                           "border-top-right-radius: 5px;"
                           "border-top-left-radius: 5px;"
                           "font: bold;}"

                           "QTabBar::tab:!selected {background-color: "
                           "#DDD6FF;"
                           "margin-top: 2px;"
                           "font: normal;}"

                           "QTabBar::tab:hover {background-color: "
                           "qlineargradient(x1:0, y1:0, x2:0, y2:1,"
                           "stop:0 #cfcccc,"
                           "stop:0.5 #333232,"
                           "stop:0.51 #000000,"
                           "stop:1 #585c5f);"
                           "color: #00cc00;"
                           "font: bold;}"
                           );
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
        //int k = QMessageBox::question(this,tr("New Updates"),tr("New Updates Available! Download?"),QMessageBox::Yes|QMessageBox::No,QMessageBox::No);
        //if(k == QMessageBox::Yes){
            QSettings iniSettings("ServerEmployee.ini",QSettings::IniFormat);
            iniSettings.setValue("Version",iniVersion);
            iniSettings.sync();

            resultUpdates = true;
            exeVersion();
        //}
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
//    QSettings settings("AO_Batrakov_Inc.", "ServerEmployee");
//    settings.setValue("QUIT",true);
    QApplication::setQuitOnLastWindowClosed(true);
    //MainWindow::close();
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
