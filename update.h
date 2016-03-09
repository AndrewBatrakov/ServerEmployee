#ifndef UPDATE_H
#define UPDATE_H

#include <QNetworkAccessManager>
#include <QtNetwork>
#include <QUrl>
#include <QtWidgets>
#include <QtGui>

QT_BEGIN_NAMESPACE
class QDialogButtonBox;
class QFile;
class QLabel;
class QLineEdit;
class QProgressDialog;
class QPushButton;
class QSslError;
class QAuthenticator;
class QNetworkReply;
QT_END_NAMESPACE

class Update : public QDialog
{
    Q_OBJECT
public:
    Update(QWidget *parent = 0);

public slots:
    void iniVersion();

private slots:
    bool newVersion();

    void exeVersion();
    void tranceVersion();

    void cancelDownLoadIni();
    void cancelDownLoadExe();
    void cancelDownLoadTrance();

    void restartProgramm();

    void httpDoneIni();
    void httpDoneExe();
    void httpDoneTrance();

    void closeConnection();

    void updateDataReadProgressExe(qint64, qint64);
    void updateDataReadProgressTrance(qint64, qint64);

    void httpReadyReadIni();
    void httpReadyReadExe();
    void httpReadyReadTrance();

private:
    QUrl url;

    QNetworkAccessManager httpIni;
    QNetworkAccessManager httpExe;
    QNetworkAccessManager httpTrance;

    QNetworkReply *replyIni;
    QNetworkReply *replyExe;
    QNetworkReply *replyTrance;

    QFile *fileHttpIni;
    QFile *fileHttpExe;
    QFile *fileHttpTrance;
    QString nowVersion;

    QProgressDialog *progressDialogIni;
    QProgressDialog *progressDialogExe;
    QProgressDialog *progressDialogTrance;
};

#endif // UPDATEHTTP_H
