#include <QtWidgets>
#include "mainwindow.h"
#include "databasedirection.h"
#include "update.h"
#include <cstdlib>

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    QTranslator translator;
    if(translator.load("ServerEmployee_ru.qm"))
        app.installTranslator(&translator);

    QApplication::setQuitOnLastWindowClosed(false);

    Update update;
    update.iniVersion();

    MainWindow mW;
    mW.show();
    
    return app.exec();
}
