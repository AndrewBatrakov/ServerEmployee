#ifndef EXPORTXML_H
#define EXPORTXML_H

#include <QAxObject>
#include <QtNetwork>

class ExportXML : public QObject
{
    Q_OBJECT

public:
    ExportXML();
    ~ExportXML();

    QAxObject *excel;
    QAxObject *queryAll;
    QAxObject *rowAll;
    QAxObject *resAll;
    QAxObject *rrr;

public slots:
    void openImport(bool);

private slots:
    QString convertStToDate(QString);

private:
    QUrl url;
    QNetworkAccessManager httpIni;
    QNetworkReply *replyIni;
};

#endif // ONECIMPORT_H
