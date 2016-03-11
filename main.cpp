#include <QCoreApplication>
#include "timeupdate.h"
#include "update.h"
#include <cstdlib>
#include "myservice.h"

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);
    //MyService service(argc, argv);

    Update upDate;
    upDate.iniVersion();

    TimeUpdate timeUpdate;

    return app.exec();
}
