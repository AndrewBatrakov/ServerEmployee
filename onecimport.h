#ifndef ONECIMPORT_H
#define ONECIMPORT_H

#include <QtWidgets/QDialog>
#include "lineedit.h"
#include <QLibrary>
#include <QAxObject>


class QAxObject;

class ImportOneForm : public QDialog
{
    Q_OBJECT

public:
    ImportOneForm(QWidget *parent = 0);
    QAxObject *excel;

public slots:
    void openImport();

private slots:
    QString convertStToDate(QString);

private:
    QAxObject *queryAll;
    QAxObject *resAll;
    QAxObject *rowAll;
};

#endif // ONECIMPORT_H
