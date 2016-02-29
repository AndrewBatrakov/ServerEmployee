#include "xmlworker.h"
#include <QAxObject>
#include <QtCore>

XMLWorker::XMLWorker ()
{
    exportXml = NULL;
}

XMLWorker::~ XMLWorker()
{
    if (exportXml != NULL) {
        delete exportXml;
    }
}

void XMLWorker::process()
{
   //CoInitializeEx(NULL, COINIT_MULTITHREADED);

    exportXml = new ExportXML;
    exportXml->openImport(false);

    emit finished();
    return ;
}

void XMLWorker::stop() {
    if(exportXml != NULL) {
        exportXml->deleteLater();
    }
    return ;
}

