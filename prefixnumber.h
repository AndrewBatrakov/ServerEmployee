#ifndef PREFIXNUMBER_H
#define PREFIXNUMBER_H

#include <QDialog>
#include "lineedit.h"

class QLabel;
class LineEdit;
class QDialogButtonBox;

class PrefixNumber : public QDialog
{
    Q_OBJECT
public:
    PrefixNumber(QWidget *parent);

private slots:
    void editPN();

private:
    QLabel *labelPrefix;
    LineEdit *editPrefix;
    QPushButton *savePushButton;
    QPushButton *cancelPushButton;
    QDialogButtonBox *buttonBox;

};

#endif // PREFIXNUMBER_H
