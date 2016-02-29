#ifndef XMLWORKER_H
#define XMLWORKER_H

#include <QObject>
#include "exportxml.h"

class XMLWorker : public QObject
{
    Q_OBJECT

private:
    ExportXML *exportXml; 		/* построитель отчетов */

public:
    XMLWorker();
    ~XMLWorker();

public slots:
    void process();
    void stop();    	/*  останавливает построитель отчетов */

signals:
    void finished(); 	/* сигнал о завершении  работы построителя отчетов */

};

#endif // XMLWORKER_H
