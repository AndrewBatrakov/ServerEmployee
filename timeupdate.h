#ifndef TIMEUPDATE_H
#define TIMEUPDATE_H

#include "update.h"
#include "exportxml.h"

class TimeUpdate : public QObject
{
    Q_OBJECT
    
public:
    TimeUpdate();
    ~TimeUpdate();

private slots:
    void updateTime();

private:
    Update updateFtp;
    ExportXML exportXML;
};

#endif // TIMEUPDATE_H
