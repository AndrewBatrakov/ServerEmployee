#ifndef PUTFTP_H
#define PUTFTP_H

#include <QtWidgets>
#include <QtNetwork>

class PutFtp : public QDialog
{
    Q_OBJECT
public:
    PutFtp(QWidget *);

public slots:
    void putFtp();

private slots:
    void updateDataTransferProgress(qint64, qint64);

private:
    QFile *filePut;
    QProgressDialog *progressDialog;
    QNetworkAccessManager *manager;
    QNetworkReply *reply;
    QUrl url;
};

#endif // PUTFTP_H
