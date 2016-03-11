#include "putftp.h"

PutFtp::PutFtp()
{
}

void PutFtp::putFile(QString nameOfFile)
{
    filePut = new QFile;
    filePut->setFileName(nameOfFile);
    filePut->open(QIODevice::ReadOnly);

    QString name = "ftp://";
    name += "91.102.219.74";
    name += "/QtProject/ServerEmployee/Change/";
    name += nameOfFile;

    QUrl url(name);
    url.setUserName("ftpsites");
    url.setPassword("Andrew123");
    url.setPort(2210);

    QEventLoop loop;
    manager = new QNetworkAccessManager(0);
    reply = manager->put(QNetworkRequest(url),filePut);
    connect(reply,SIGNAL(finished()),&loop,SLOT(quit()));
    loop.exec();
}
