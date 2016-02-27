#ifndef FTPFORM_H
#define FTPFORM_H

#include <QtNetwork>
#include <QtWidgets>
#include <QtGui>
#include "lineedit.h"

class FtpForm : public QDialog
{
    Q_OBJECT

public:
    FtpForm(QWidget *parent = 0);
    void done(int);

public slots:
    void saveFtpSettings();
    void replyFinished(QNetworkReply *fin);
    void conn();

private:
    QLabel *nameHost;
    LineEdit *editHost;
    QLabel *nameLogin;
    LineEdit *editLogin;
    QLabel *namePass;
    LineEdit *editPass;

    QPushButton *saveButton;
    QPushButton *cancelButton;
    QDialogButtonBox *buttonBox;

    //QFtp *ftpPut;
    QFile *filePut;
};

#endif // FTPFORM_H
