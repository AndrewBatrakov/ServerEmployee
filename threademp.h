#ifndef THREADEMP_H
#define THREADEMP_H

#include <QThread>

class ThreadEmp : public QThread
{
    Q_OBJECT
public:
    ThreadEmp();
    void stop();

protected:
    void run();

private:
    volatile bool stopped;
};

#endif // THREADEMP_H
