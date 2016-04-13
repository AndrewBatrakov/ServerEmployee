#ifndef PUTFTP_H
#define PUTFTP_H

#include <QtNetwork>

class PutFtp : public QObject
{
    Q_OBJECT
public:
    PutFtp();
    ~PutFtp();

public slots:
    void putFile(QString);

private:
    QFile *filePut;
    QNetworkAccessManager *manager;
    QNetworkReply *reply;
    QUrl url;
};

#endif // PUTFTP_H
