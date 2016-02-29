#ifndef SESSION_H
#define SESSION_H

#include <QObject>

class Session : public QObject
{
    Q_OBJECT

    public:
        Session(QObject *parent);
        ~Session();

        void buildReports();

    private:
        void addThread();
        void stopThreads();

    signals:
        void stopAll(); //остановка всех потоков
};

#endif // SESSION_H
