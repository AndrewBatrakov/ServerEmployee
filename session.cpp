#include "session.h"
#include "xmlworker.h"

Session::Session(QObject *parent) : QObject(parent)
{
}

void Session::addThread()
{
    XMLWorker *worker = new XMLWorker;
    QThread *thread = new QThread;
    worker->moveToThread(thread);


/*  Теперь внимательно следите за руками.  Раз: */
    connect(thread, SIGNAL(started()), worker, SLOT(process()));
/* … и при запуске потока будет вызван метод process(), который создаст построитель отчетов, который будет работать в новом потоке

Два: */
    connect(worker, SIGNAL(finished()), thread, SLOT(quit()));
/* … и при завершении работы построителя отчетов, обертка построителя передаст потоку сигнал finished() , вызвав срабатывание слота quit()

Три:
*/
    connect(this, SIGNAL(stopAll()), worker, SLOT(stop()));
/* … и Session может отправить сигнал о срочном завершении работы обертке построителя, а она уже остановит построитель и направит сигнал finished() потоку

Четыре: */
    connect(worker, SIGNAL(finished()), worker, SLOT(deleteLater()));
/* … и обертка пометит себя для удаления при окончании построения отчета

Пять: */
    connect(thread, SIGNAL(finished()), thread, SLOT(deleteLater()));
/* … и поток пометит себя для удаления, по окончании построения отчета. Удаление будет произведено только после полной остановки потока.

И наконец :
*/
    thread->start();
/* Запускаем поток, он запускает RBWorker::process(), который создает ReportBuilder и запускает  построение отчета */

    return ;
}


void Session::stopThreads()  /* принудительная остановка всех потоков */
{
    emit  stopAll();
/* каждый RBWorker получит сигнал остановиться, остановит свой построитель отчетов и вызовет слот quit() своего потока */
}

void Session::buildReports()
{
    //CoInitializeEx(NULL, COINIT_MULTITHREADED);
    addThread();
    return ;
}

Session::~Session()
{
    stopThreads();  /* останавливаем и удаляем потоки  при окончании работы сессии */
}
