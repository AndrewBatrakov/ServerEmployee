#include <QtGui>
#include "databasedirection.h"

DataBaseDirection::DataBaseDirection()
{
}

bool DataBaseDirection::connectDataBase()
{
    bool okBool;
    QSqlDatabase globalDataBaseConnection = QSqlDatabase::addDatabase("QMYSQL");
    globalDataBaseConnection.setDatabaseName("traficsafety");
    globalDataBaseConnection.setHostName("91.102.219.74");
    globalDataBaseConnection.setUserName("root");
    globalDataBaseConnection.setPassword("Helga2210");
    if(!globalDataBaseConnection.open()){
        QMessageBox::warning(0, QObject::tr("MySQL Database Error"),
                             globalDataBaseConnection.lastError().text());
        okBool = false;
    }else{
        okBool = true;
    }
    return okBool;
}

void DataBaseDirection::removeDateBase()
{
    QSqlDatabase db = QSqlDatabase::database();
    db.close();
    //db.removeDatabase();
}
