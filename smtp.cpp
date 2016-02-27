#include "smtp.h"

SMTP::SMTP(QString host, QString userName, QString pass)
{
    qDebug()<<"---Smtp start";
    this->smtphost = smtphost;
    this->smtpusername = smtpusername;
    this->smtppass = smtppass;
}

void SMTP::SetHost(QString host, QString userName, QString pass)
{
    this->smtphost = host;
    this->smtpusername = userName;
    this->smtppass = pass;
}

void SMTP::Send(const QString &from, const QString &to, const QString &subject, const QString &body)
{
    int waitTime = 5000;
    this->from = from;
    rcpt = to;
    errorMSG.clear();
    timeOut = waitTime;
    lineSend = 0;
    isConnected = false;
    message = timeStampMail() + "\n";
    message.append("User-Agent: Qt SMTP client\n");
    message.append("X-Accept-Language: en-us, en\n");
    message.append("MIME-Version: 1.0\n");
    message.append("To: " + to + "\n");
    message.append("From: " + from + " <" + from + ">\n");
    message.append("Subject: " + subject + "\n");

    smtpSocket = new QTcpSocket(this);
    connect(this,SIGNAL(connectorSuccess()),this,SLOT(readLiner()));
    connect(this,SIGNAL(sendLine()),this,SLOT(putSendLine()));
    if(smtphost.size() > 0)
        smtpSocket->connectToHost(smtphost, 25);
    else
        smtpSocket->connectToHost("localhost",25);

    if(smtpSocket->waitForConnected(timeOut))
    {
        qDebug()<<"--- connected on "<<smtphost;
        if(smtpSocket->waitForReadyRead(timeOut))
        {
            isConnected = true;
            emit connectorSuccess();
        }
    }else{
        emit errorCloseAll();
    }
}

void SMTP::readLiner()
{
    if(isConnected)
    {
        QTextCodec *codecX;
        codecX = QTextCodec::codecForMib(106);
        t = new QTextStream(smtpSocket);
        t->setCodec(codecX);

        int loops;
        while(!t->atEnd()){
            loops++;
            response = t->readLine();
        }
        if(response.size() > 0)
        {
            remoteServerName = response;
            mailStatus = response.left(3);
            if(mailStatus == "220"){
                response = "";
                lineSend = 1;
                emit sendLine();
            }
        }else{
            emit errorCloseAll();
        }
    }
}

SMTP::~SMTP()
{

}

void SMTP::putSendLine()
{
    int current = lineSend;
    qDebug()<<"Send line"<<lineSend;
    switch(current)
    {
    case 1:
        response = sendLineAndGrab("localhost");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 2;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 2:
        response =sendLineAndGrab("AUTH LOGIN");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 3;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 3:
        response =sendLineAndGrab("AUTH LOGIN");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 4;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 4:
        response =sendLineAndGrab("AUTH LOGIN");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 5;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 5:
        response =sendLineAndGrab("AUTH LOGIN");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 6;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 6:
        response =sendLineAndGrab("AUTH LOGIN");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 7;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 7:
        response =sendLineAndGrab("AUTH LOGIN");
        if(response.size() > 0){
            errorMSG.append(response);
            lineSend = 8;
            emit sendLine();
        }else{
            qDebug()<<"connection lost";
        }
        response = "";
        break;
    case 8:
        sendLineAndGrab("QUIT");
        break;
    default:
        emit errorCloseAll();
        return;
        break;
    }
}

QString SMTP::sendLineAndGrab(QString sendData)
{
    QString incommingData = "";
    if(isConnected){
        int current = lineSend;
        int loops = 0;
        qDebug()<<"Send"<<current<<"Now =>"<<sendData;
        *t << sendData << "\r\n";
        t->flush();
        if(sendData != "QUIT"){
            if(smtpSocket->waitForReadyRead(timeOut)){
                while(!t->atEnd()){
                    loops++;
                    QString opera = t->readLine()+"\n";
                    incommingData = opera + incommingData;
                    qDebug()<<loops<<"|#"<<opera<<"#|";
                }
            }else{
                delete smtpSocket;
                delete t;
                isConnected = false;
                emit successQuit();
                return incommingData;
            }
        }else{
            emit errorCloseAll();
        }
    }
    return incommingData;
}

void SMTP::disconnected()
{
    qDebug()<<"!!! disconnected";
    qDebug()<<"Error: "<<smtpSocket->errorString();
}

void SMTP::connected()
{
    output.append("connected");
    qDebug()<<"Connected ";
}

QString SMTP::encodeBase64(QString xml)
{
    QByteArray text;
    text.append(xml);
    return text.toBase64();
}

QString SMTP::decodeBase64(QString xml)
{
    QByteArray xcode("");
    xcode.append(xml);
    QByteArray precode(QByteArray::fromBase64(xcode));
    QString nodetxt = precode.data();
    return nodetxt;
}

QStringList SMTP::getErrorList()
{
    return errorMSG;
}

int SMTP::dateStamp(QString from, uint unixtime)
{
    QDateTime fromunix;
    fromunix.setTime_t(unixtime);
    QString numeric = fromunix.toString((const QString) from);
    bool ok;
    return (int)numeric.toFloat(&ok);
}

QString SMTP::timeStampMail()
{
    QDateTime timer1(QDateTime::currentDateTime());
    uint unixtime = timer1.toTime_t();
    QDateTime fromunix;
    fromunix.setTime_t(unixtime);
    return "";
}
