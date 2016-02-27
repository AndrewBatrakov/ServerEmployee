#include <QtGui>

#include "ftpform.h"

FtpForm::FtpForm(QWidget *parent) :
    QDialog(parent)
{
    nameHost = new QLabel("Host:");
    editHost = new LineEdit;

    nameLogin = new QLabel("Login:");
    editLogin = new LineEdit;

    namePass = new QLabel("Pass:");
    editPass = new LineEdit;

    QVBoxLayout *leftLayout = new QVBoxLayout;
    leftLayout->addWidget(nameHost);
    leftLayout->addWidget(nameLogin);
    leftLayout->addWidget(namePass);

    QVBoxLayout *rightLayout = new QVBoxLayout;
    rightLayout->addWidget(editHost);
    rightLayout->addWidget(editLogin);
    rightLayout->addWidget(editPass);

    QHBoxLayout *bothLayout = new QHBoxLayout;
    bothLayout->addLayout(leftLayout);
    bothLayout->addLayout(rightLayout);

    saveButton = new QPushButton(tr("Save"));
    connect(saveButton,SIGNAL(clicked()),this,SLOT(saveFtpSettings()));
    saveButton->setToolTip(tr("Save And Close Button"));

    cancelButton = new QPushButton(tr("Cancel"));
    cancelButton->setDefault(true);
    cancelButton->setStyleSheet("QPushButton:hover {color: red}");
    connect(cancelButton,SIGNAL(clicked()),this,SLOT(accept()));
    cancelButton->setToolTip(tr("Cancel Button"));

    buttonBox = new QDialogButtonBox;
    buttonBox->addButton(cancelButton,QDialogButtonBox::ActionRole);
    buttonBox->addButton(saveButton,QDialogButtonBox::ActionRole);

    QSettings settings("AO_Batrakov_Inc.", "ServerEmployee");
    editHost->setText(settings.value("Host").toString());
    editLogin->setText(settings.value("Login").toString());
    editPass->setText(settings.value("Pass").toString());

    QVBoxLayout *mainLayout = new QVBoxLayout;
    mainLayout->addLayout(bothLayout);
    //if(!onlyForRead){
        mainLayout->addWidget(buttonBox);
    //}
    setLayout(mainLayout);

    setWindowTitle(tr("Ftp settings"));
}

void FtpForm::saveFtpSettings()
{
    QSettings settings("AO_Batrakov_Inc.", "ServerEmployee");
    settings.setValue("Host", editHost->text());
    settings.setValue("Login", editLogin->text());
    settings.setValue("Pass", editPass->text());
    emit accept();
}

void FtpForm::done(int result)
{
    QDialog::done(result);
}

void FtpForm::replyFinished(QNetworkReply *reply)
{

}

void FtpForm::conn()
{

}

