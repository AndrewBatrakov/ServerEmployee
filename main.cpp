#include <QApplication>
#include "timeupdate.h"
#include "update.h"
#include <cstdlib>
#include "qaxobject.h"

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    Update upDate;
    upDate.iniVersion();

    TimeUpdate timeUpdate;

    return app.exec();
}
