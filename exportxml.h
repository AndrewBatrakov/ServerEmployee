#ifndef EXPORTXML_H
#define EXPORTXML_H

#include <QAxObject>
#include <QtNetwork>

class ExportXML : public QObject
{
    Q_OBJECT

public:
    ExportXML(QObject *parent = 0);

    QAxObject *excel;
    QAxObject *queryAll;
    QAxObject *rowAll;
    QAxObject *resAll;

public slots:
    void openImport(bool);

private slots:
    QString convertStToDate(QString);
};

#endif // ONECIMPORT_H
