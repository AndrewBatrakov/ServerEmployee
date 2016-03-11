#ifndef UPDATE_H
#define UPDATE_H

#include <QtNetwork>
#include <QtCore>

class Update : public QObject
{
    Q_OBJECT
public:
    Update();

public slots:
    void iniVersion();

private slots:
    bool newVersion();

    void exeVersion();

    void restartProgramm();

    void httpDoneIni();
    void httpDoneExe();

    void closeConnection();

    void httpReadyReadIni();
    void httpReadyReadExe();

private:
    QUrl url;

    QNetworkAccessManager httpIni;
    QNetworkAccessManager httpExe;

    QNetworkReply *replyIni;
    QNetworkReply *replyExe;

    QFile *fileHttpIni;
    QFile *fileHttpExe;

    QString nowVersion;
};

#endif // UPDATEHTTP_H
