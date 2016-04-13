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
};

#endif // ONECIMPORT_H
