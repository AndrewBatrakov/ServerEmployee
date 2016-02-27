#include "mainwindow.h"
#include "ftpform.h"
#include "threademp.h"
#include "exportxml.h"
#include "threademp.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    QWidget::setWindowFlags(Qt::Window | Qt::FramelessWindowHint | Qt::WindowStaysOnTopHint);
    startSettings();
    QAction *showHide = new QAction(tr("Show/Hide Application Window"),this);
    connect(showHide,SIGNAL(triggered()),this,SLOT(slotShowHide()));

    QAction *ftpForm = new QAction(tr("FTP settings"),this);
    connect(ftpForm,SIGNAL(triggered()),this,SLOT(ftpForm()));

    QAction *exitAction = new QAction(tr("Exit Application"),this);
    connect(exitAction,SIGNAL(triggered()),qApp,SLOT(quit()));

    iconMenu = new QMenu(this);
    iconMenu->addAction(showHide);
    iconMenu->addAction(ftpForm);
    iconMenu->addAction(exitAction);

    systemTryIcon = new QSystemTrayIcon(this);
    systemTryIcon->setIcon(QPixmap(":/sunflower.png"));
    systemTryIcon->setContextMenu(iconMenu);
    systemTryIcon->show();

    QPixmap pixWater(":/sunflower.png");
    setWindowIcon(pixWater);
    setWindowTitle(tr("Server Employee"));

    timeLabel = new QLCDNumber(8);

    QSplitter *splitter = new QSplitter(Qt::Horizontal);
    splitter->setFrameStyle(QFrame::StyledPanel);
    //splitter->addWidget(leftPanel);
    splitter->addWidget(timeLabel);
    timeLabel->display(QTime::currentTime().toString("HH:mm:ss"));
    splitter->setStretchFactor(1,1);
    setCentralWidget(splitter);

    QTimer *timer = new QTimer(this);
    connect(timer,SIGNAL(timeout()),this,SLOT(updateTime()));
    timer->start(1000);

    QFile file(":/QWidgetStyle.txt");
    file.open(QFile::ReadOnly);
    QString styleSheetString = QLatin1String(file.readAll());
    QWidget::setStyleSheet(styleSheetString);

    slotShowHide();
    //update.iniVersion();
    ThreadEmp threadEmp;
    threadEmp.start();
}

MainWindow::~MainWindow()
{
    delete systemTryIcon;
}

void MainWindow::closeEvent(QCloseEvent *)
{

}

void MainWindow::slotShowHide()
{
    setVisible(!isVisible());
}

void MainWindow::slotShowMessage(QString tryIconString)
{
    systemTryIcon->showMessage(tr("Write Employee:"),tryIconString,QSystemTrayIcon::Information,1000);
}

void MainWindow::startProcedure()
{

}

void MainWindow::updateTime()
{

    timeLabel->display(QTime::currentTime().toString("HH:mm:ss"));
    QString currTime = QTime::currentTime().toString();
    if(currTime == "00:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
        //ExportXML exportXML;
        //exportXML.openImport();
    }else if(currTime == "03:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }else if(currTime == "06:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }else if(currTime == "09:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }else if(currTime == "12:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }else if(currTime == "15:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }else if(currTime == "18:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }else if(currTime == "21:00:00"){
        timeLabel->setStyleSheet("QLCDNumber {"
                                 "color: darkblue;"
                                 "background-color: #FFF951;}");
        ThreadEmp threadEmp;
        threadEmp.start();
    }

    if(currTime == "23:30:00"){
        update.iniVersion();
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

    timeLabel->setStyleSheet("QLCDNumber {"
                             "color: darkblue;"
                             "background-color: #08FF52;}");
}

void MainWindow::startSettings()
{
    QSettings settings("AO_Batrakov_Inc.", "ServerEmployee");
    QPoint pos = settings.value("pos").toPoint();
    QSize size = settings.value("size").toSize();
    if(pos.isNull()){
        pos = QPoint(qApp->desktop()->width() - 165,0);
        size.setHeight(80);
        size.setWidth(160);
    }

    resize(size);
    move(pos);
}

void MainWindow::ftpForm()
{
    FtpForm ftpForm(this);
    ftpForm.exec();
}

