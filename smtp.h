#ifndef SMTP_H
#define SMTP_H

#include <QObject>
#include <QtNetwork>
#include <QDebug>
//#include <QtTest/QtTest>

class SMTP : public QObject
{
    Q_OBJECT

public:
    SMTP(QString host, QString userName, QString pass);
    ~SMTP();
    void SetHost(QString host, QString userName, QString pass);
    void Send(const QString &from, const QString &to, const QString &subject, const QString &body);
    int lineSend;
    QStringList errorMSG;
    QStringList getErrorList();
    
signals:
    void status(const QString &);
    void connectorSuccess();
    void sendLine();
    void errorCloseAll();
    void successQuit();
    
public slots:
    void disconnected();
    void connected();
    void readLiner();
    void putSendLine();

private:
    bool isConnected;
    QString smtphost;
    QString smtpusername;
    QString smtppass;
    QString message;
    QString output;
    QString remoteServerName;
    QString mailStatus;
    QTextStream *t;
    QTcpSocket *smtpSocket;
    QString from;
    QString rcpt;
    QString response;
    QString sendLineAndGrab(QString sendData);
    int timeOut;
    QString encodeBase64(QString xml);
    QString decodeBase64(QString xml);
    int dateStamp(QString from, uint unixtime);
    QString timeStampMail();
    
};

#endif // SMTP_H
