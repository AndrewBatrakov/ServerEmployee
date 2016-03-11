#ifndef TIMEUPDATE_H
#define TIMEUPDATE_H

#include "update.h"

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
};

#endif // TIMEUPDATE_H
