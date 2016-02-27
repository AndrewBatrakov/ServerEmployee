#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QtWidgets>
#include "update.h"
#include "update.h"
//#include "onecimport.h"
//#include "exportxml.h"

class MainWindow : public QMainWindow
{
    Q_OBJECT
    
public:
    MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void slotShowMessage(QString tryIconString);
    void startProcedure();

private slots:
    void slotShowHide();
    void closeEvent(QCloseEvent *);
    void updateTime();
    void startSettings();
    void ftpForm();

private:
    QSystemTrayIcon *systemTryIcon;
    QMenu *iconMenu;

    QPushButton *startButton;
    QLCDNumber *timeLabel;
    Update updateFtp;
    Update update;
    //ImportOneForm importOneC;
    //ExportXML exportXML;
};

#endif // MAINWINDOW_H
