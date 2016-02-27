#ifndef EXPORTXML_H
#define EXPORTXML_H

#include <QDialog>
#include "lineedit.h"
#include <QLibrary>
#include <QAxObject>
#include <QtNetwork>

class QLabel;
class LineEdit;
class QDialogButtonBox;
class QComboBox;
class QDateEdit;
class QTableView;
class QAxObject;

class ExportXML : public QDialog
{
    Q_OBJECT

public:
    ExportXML(QWidget *parent = 0);

    QAxObject *excel;
    QAxObject *queryAll;
    QAxObject *rowAll;
    QAxObject *resAll;

public slots:
    void openImport();

private slots:
    QString convertStToDate(QString);

private:
    //QFtp *ftpPut;
    //QFile *filePut;
    //QLabel *nameCount;
    //QLabel *editCount;
    //QLabel *nameAll;
    //QLabel *editAll;

};

#endif // ONECIMPORT_H
