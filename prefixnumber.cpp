#include <QtGui>
#include <QtSql>
#include "prefixnumber.h"

PrefixNumber::PrefixNumber(QWidget *parent) :
    QDialog(parent)
{
    labelPrefix = new QLabel(tr("Name:"));
    editPrefix = new LineEdit;

    QRegExp regExpFamiliya("[A-Za-z]{3}");
    editPrefix->setValidator(new QRegExpValidator(regExpFamiliya,this));
    labelPrefix->setBuddy(editPrefix);

    savePushButton = new QPushButton(tr("Save"));
    connect(savePushButton,SIGNAL(clicked()),this,SLOT(editPN()));

    cancelPushButton = new QPushButton(tr("Cancel"));
    cancelPushButton->setDefault(true);
    cancelPushButton->setStyleSheet("QPushButton:hover {color: red}");
    connect(cancelPushButton,SIGNAL(clicked()),this,SLOT(accept()));

    buttonBox = new QDialogButtonBox;
    buttonBox->addButton(cancelPushButton,QDialogButtonBox::ActionRole);
    buttonBox->addButton(savePushButton,QDialogButtonBox::ActionRole);

    QSettings settings("AO_Batrakov_Inc.", "ServerEmployee");
    if(settings.value("numprefix").toString() == ""){
        editPrefix->setText(tr("Three letters of the Latin alphabet"));
    }else{
        editPrefix->setText(settings.value("numprefix").toString());
    }
    editPrefix->setReadOnly(false);
    QGridLayout *mainLayout = new QGridLayout;
    mainLayout->addWidget(labelPrefix,0,0);
    mainLayout->addWidget(editPrefix,0,1);
    mainLayout->addWidget(buttonBox,1,1);
    setLayout(mainLayout);

    setWindowTitle(tr("Prefix Number "));
}

void PrefixNumber::editPN()
{
    QSettings settings("AO_Batrakov_Inc.", "ServerEmployee");
    editPrefix->text().toUpper();
    settings.setValue("numprefix", editPrefix->text());

    close();
}
