#ifndef DATABASEDIRECTION_H
#define DATABASEDIRECTION_H

#include <QtSql>
#include <QtWidgets>
#include <QDialog>

class DataBaseDirection : public QDialog
{
    Q_OBJECT
public:
    DataBaseDirection();

public:
    bool connectDataBase();
    void removeDateBase();

};
#endif // DATABASEDIRECTION_H
