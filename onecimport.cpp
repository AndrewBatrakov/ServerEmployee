#include <QtGui>
#include <QtWidgets>
#include <QtSql>

#include "onecimport.h"
#include "numprefix.h"
#include "mainwindow.h"
#include <qaxobject.h>
#include "databasedirection.h"

ImportOneForm::ImportOneForm(QWidget *parent)
    :QDialog(parent)
{
    queryAll = new QAxObject;
    rowAll = new QAxObject;
    resAll = new QAxObject;
}

void ImportOneForm::openImport()
{
    DataBaseDirection direc;
    direc.connectDataBase();

    //QTextCodec::setCodecForCStrings(QTextCodec::codecForName("UTF8"));
    QTextCodec *codecTr = QTextCodec::codecForName("Windows-1251");
    QTextDecoder *decoder = codecTr->makeDecoder();
    //QTextCodec::setCodecForTr(QTextCodec::codecForName("Windows-1251"));
    excel = new QAxObject("v83.ComConnector",this);

    if(!excel->isNull()){

        QString queryTextFizLiz, queryTextKarty, queryTextDolzn, queryTextObrazovanie, queryTextRodstva, queryTextStaz,
                queryTextPhone, queryTextMedosmotr, queryTextPromBezop, queryTextOhranaTruda, queryTextPTM, queryTextRost,
                queryTextProfessii, queryTextPassport;

//        QByteArray text = ("������� "
//                           "��������������.��������������� ��� DeletionMark, "
//                           "��������������.��� ��� Kod, "
//                           "��������������.������������ ��� FIO, "
//                           "��������������.������������ ��� DataRozdeniya "
//                           "�� ����������.�������������� ��� �������������� "
//                           "��� "
//                           "��������������.��������� <> ������ "
//                           "����������� �� "
//                           "FIO "
//                           "������������������");
//        queryTextFizLiz = codecTr->toUnicode(text);

        QByteArray text = ("������� "
                           "��������������.��� ��� Kod, "
                           "��������������.������������ ��� FIO, "
                           "��������������.������������ ��� DataRozdeniya, "
                           "��������������.��������������� ��� DeletionMark, "
                           "���������������������������������.���������.���, "
                           "���������������������������������.�������������������������.������������ ��� Obosobl, "
                           "���������������������������������.������������������������.������������ ��� Subdivision, "
                           "���������������������������������.���������.������������ ��� Post, "
                           "���������������������������������.������������.������������ ��� GrafikRabot, "
                           "���������������������������������.���������.������������������ ��� DatePostupleniya, "
                           "���������������������������������.���������.�������������� ��� DateUvolneniya, "
                           "���_���_�����������������������.�������� ��� KodKarty "
                           "�� "
                           "����������.�������������� ��� �������������� "
                           "����� ���������� ���������������.��������������������.������������� ��� ��������������������������������� "
                           "�� ��������������.������ = ���������������������������������.���������.�������.������ "
                           "����� ���������� ���������������.���_���_����������.������������� ��� ���_���_����������������������� "
                           "�� ��������������.������ = ���_���_�����������������������.�������.������ "
                           "��� "
                           "�� ��������������.��������� "
                           "� �� ��������������.��������������� "
                           "� ���������������������������������.���������.�������������� = ���������(1, 1, 1) "
                           "� �� ���������������������������������.���������.������������������ = ���������(1, 1, 1) "
                           "� ���_���_�����������������������.������� "
                           "����������� �� "
                           "FIO");
        queryTextFizLiz = codecTr->toUnicode(text);

//        text = ("������� ������ 1 "
//                "���_���_�����������������������.�������� ��� KodKarty "
//                "�� "
//                "���������������.���_���_����������.������������� ��� "
//                "���_���_����������������������� "
//                "��� ���_���_�����������������������.�������.��� = \"%1\" "
//                "����������� �� "
//                "���_���_�����������������������.������ ����");

//        //"� ���_���_�����������������������.������� = ������ "
//        queryTextKarty = codecTr->toUnicode(text);

//        text = ("������� "
//                "���������������������������������.������������������������.������������ ��� Subdivision, "
//                "���������������������������������.���������.������������ ��� Post, "
//                "���������������������������������.���������.Code ��� TabelNumber, "
//                "���������������������������������.���������.������������������ ��� DatePostupleniya, "
//                "���������������������������������.���������.�������������� ��� DateUvolneniya, "
//                "���������������������������������.���������.������������.������������ ��� GrafikRabot, "
//                "���������������������������������.�������������������������.������������ ��� Obosobl "
//                "�� "
//                "���������������.��������������������.������������� ��� ��������������������������������� "
//                "��� ���������������������������������.���������.�������.��� = \"%1\"");
//        queryTextDolzn = codecTr->toUnicode(text);

        //Passport
        text = ("������� "
                "�����������������������������������.������������� ��� PassportSeriya, "
                "�����������������������������������.������������� ��� PassportNomer, "
                "�����������������������������������.������������������ ��� PassportDataVidachi, "
                "�����������������������������������.���������������� ��� KemVidan "
                "�� "
                "���������������.����������������������.������������� ��� ����������������������������������� "
                "��� "
                "�����������������������������������.�������.��� = \"%1\"");
        queryTextPassport = codecTr->toUnicode(text);

        text = ("������� "
                "�������������������������.��������������.������������ ��� VidObrazovaniya, "
                "�������������������������.����������������.������������ ��� UchebnoeZavedenie, "
                "�������������������������.�������������.������������ ��� Spezialnost, "
                "�������������������������.������ ��� Diplom, "
                "�������������������������.������������ ��� GodOconchaniya, "
                "�������������������������.������������ ��� Kvalificaziya "
                "�� "
                "����������.��������������.����������� ��� ������������������������� "
                "��� "
                "�������������������������.������.��� = \"%1\"");
        queryTextObrazovanie = codecTr->toUnicode(text);

        text = ("������� "
                "�������������������������.��������������.������������ ��� StepenRodstva, "
                "�������������������������.��� ��� NameRodstva, "
                "�������������������������.������������ ��� BirthRodstva "
                "�� "
                "����������.��������������.����������� ��� ������������������������� "
                "��� "
                "�������������������������.������.��� = \"%1\"");
        queryTextRodstva = codecTr->toUnicode(text);

        text = ("������� "
                "��������������������.������������� ��� Tephone "
                "�� "
                "���������������.�������������������� ��� �������������������� "
                "��� "
                "��������������������.������.��� = \"%1\"");
        queryTextPhone = codecTr->toUnicode(text);

        text = ("������� ������ 1 "
                "���_���_�������������������������������.������������ ��� NomerSpravki, "
                "���_���_�������������������������������.����� ��� Goden, "
                "���_���_�������������������������������.��������������� ��� DataProhojdeniya, "
                "���_���_�������������������������������.����� ��� NomerDocumenta, "
                "���_���_�������������������������������.���������������� ��� Proyden, "
                "���_���_�������������������������������.���� ��� Data "
                "�� "
                "��������.���_���_������������������������������� ��� ���_���_������������������������������� "
                "��� "
                "���_���_�������������������������������.���������.��� = \"%1\" "
                "����������� �� "
                "���_���_�������������������������������.���� ����");
        queryTextMedosmotr = codecTr->toUnicode(text);

        text = ("������� ������ 1 "
                "���_���_���������������������������������������������������������������.�����������������.������������ ��� ProgrammaObucheniya, "
                "���_���_����������������������������������������������.����� ��� Nomer, "
                "���_���_����������������������������������������������.���� ��� Data, "
                "���_���_����������������������������������������������.���������� ��� ObChasov "
                "�� "
                "��������.���_���_����������������������������������������������.����������������� ��� ���_���_��������������������������������������������������������������� "
                "����� ���������� ��������.���_���_���������������������������������������������� ��� ���_���_���������������������������������������������� "
                "�� ���_���_���������������������������������������������������������������.������ = ���_���_����������������������������������������������.������ "
                "��� "
                "���_���_����������������������������������������������.���������.���������.�������.��� = \"%1\"");
        queryTextPromBezop = codecTr->toUnicode(text);

        text = ("������� ������ 1 "
                "���_���_�������������������������������������������������.���� ��� DataUdostovereniya, "
                "���_���_�������������������������������������������������.��������.���� ��� ProtokolData, "
                "���_���_�������������������������������������������������.��������.�����������������.������������ ��� ProtocolProgrammaObucheniya, "
                "���_���_�������������������������������������������������.��������.����� ��� ProtocolNomer, "
                "���_���_�������������������������������������������������.��������.���������� ��� ProtokolObChasov "
                "�� "
                "��������.���_���_������������������������������������������������� ��� ���_���_������������������������������������������������� "
                "��� "
                "���_���_�������������������������������������������������.���������.��� = \"%1\" "
                "����������� �� "
                "���_���_�������������������������������������������������.���� ����");
        queryTextOhranaTruda = codecTr->toUnicode(text);

        text = ("������� ������ 1 "
                "���_���_��������������������������������.���� ��� Data, "
                "���_���_��������������������������������.����� ��� Number, "
                "���_���_�������������������������������������������������.����� ��� PTMNumber, "
                "���_���_�������������������������������������������������.���� ��� PTMDate, "
                "���_���_�������������������������������������������������.���������� ��� PTMObChasov, "
                "���_���_�������������������������������������������������.���������������������� ��� PTMSostav, "
                "���_���_�������������������������������������������������.�����������������.������������ ��� PTMProgramma "
                "�� "
                "��������.���_���_�������������������������������� ��� ���_���_�������������������������������� "
                "����� ���������� ��������.���_���_������������������������������������������������� ��� ���_���_������������������������������������������������� "
                "�� ���_���_��������������������������������.�������� = ���_���_�������������������������������������������������.������ "
                "��� "
                "���_���_��������������������������������.���������.��� = \"%1\" "
                "����������� �� "
                "���_���_��������������������������������.���� ����");
        queryTextPTM = codecTr->toUnicode(text);

        text = ("������� "
                  "���_���_������������������������.���� ��� Rost, "
                  "���_���_������������������������.������������ ��� OdezhdaZim, "
                  "���_���_������������������������.������������ ��� OdezhdaLet, "
                  "���_���_������������������������.����������� ��� ObuvZim, "
                  "���_���_������������������������.����������� ��� ObuvLet, "
                  "���_���_������������������������.������������ ��� GolUbor "
                  "�� "
                  "����������.���_���_������������������������ ��� ���_���_������������������������ "
                  "��� "
                  "���_���_������������������������.��������.�������.��� = \"%1\"");
        queryTextRost = codecTr->toUnicode(text);

        text = ("������� "
                "�������������������.��������.������������, "
                "�������������������.�����������, "
                "�������������������.�������������, "
                "�������������������.���������� "
                "�� "
                "����������.��������������.����� ��� ������������������� "
                "��� "
                "�������������������.������.��� = \"%1\"");
        queryTextStaz = codecTr->toUnicode(text);

        text = ("������� "
                  "�����������������������.��������� "
                  "�� "
                  "����������.��������������.��������� ��� ����������������������� "
                  "��� "
                  "�����������������������.������.��� = \"%1\"");
        queryTextProfessii = codecTr->toUnicode(text);

        QAxObject *rrr = excel->querySubObject("Connect(QVariant &)", QVariant("Srvr=\"rif-1c\";Ref=\"1c_zupkorp\";Usr=\"ManilchukNM\";Pwd=\"1Q34WER\""));
        /*if(rrr->isNull()){
        QMessageBox::warning(this,"Connect","Don't connect to 1C!!!");
    }else{
        QMessageBox::warning(this,"Connect","Connect to 1C!!!");
    }*/

        QAxObject *queryFizLiz = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
        queryFizLiz->dynamicCall("Text", queryTextFizLiz);
        QAxObject *resFizLiz = queryFizLiz->querySubObject("Execute()");
        QAxObject *rowFizLiz = resFizLiz->querySubObject("Choose()");

        QProgressDialog progressR(this);

        progressR.setWindowModality(Qt::WindowModal);
        progressR.setMinimum(1);
        int cou = rowFizLiz->dynamicCall("Count()").toInt();
        progressR.setMaximum(cou);
        progressR.activateWindow();

        QString personName,dateName,kodSsylka,kodKarty,subId, postId, orgId, grafikRabotId;
        QString subName,postName,tabNumber,grafikRabot,tephone;
        QString personId, passportSeriya, passportNomer, passportDataVidachi, kemVidan;
        QString vidObrazovaniya,uchebnoeZavedenie,spezialnost,diplom,godOconchaniya,kvalificaziya;
        QString stepenRodstva,nameRodstva,dateRodstva;

        QDate datePostupleniya,dateUvolneniya;
        bool deletionMark;
        int lll = 0;

        while(rowFizLiz->dynamicCall("Next()").toBool()){
            if(progressR.wasCanceled())
                break;

            progressR.setWindowTitle(QString::number(lll));
            progressR.setValue(lll);
            progressR.show();
            progressR.activateWindow();

            personName = rowFizLiz->dynamicCall("FIO").toString().simplified();
            dateName = rowFizLiz->dynamicCall("DataRozdeniya").toString();
            kodSsylka = rowFizLiz->dynamicCall("Kod").toString();
            deletionMark = rowFizLiz->dynamicCall("DeletionMark").toBool();
            kodKarty = rowFizLiz->dynamicCall("KodKarty").toString().simplified();
            subName = rowFizLiz->dynamicCall("Subdivision").toString().simplified();
            postName = rowFizLiz->dynamicCall("Post").toString().simplified();
            tabNumber = rowFizLiz->dynamicCall("TabelNumber").toString();
            datePostupleniya = rowFizLiz->dynamicCall("DatePostupleniya").toDate();
            dateUvolneniya = rowFizLiz->dynamicCall("DateUvolneniya").toDate();
            grafikRabot = rowFizLiz->dynamicCall("GrafikRabot").toString().simplified();

            if(deletionMark){
                continue;
            }

            if(dateName != "100-01-01T00:00:00"){
                QSqlQuery queryP;
                queryP.prepare("SELECT * FROM employee "
                                 "WHERE employeename = :personname AND "
                                 "datebirthday = :datebirth;");
                queryP.bindValue(":personname",personName);
                queryP.bindValue(":datebirth",dateName);
                queryP.exec();
                queryP.next();
                if(queryP.isValid()){
                    ++lll;
                    continue;
                }
//                queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
//                queryAll->dynamicCall("Text", queryTextKarty.arg(kodSsylka));
//                resAll = queryAll->querySubObject("Execute()");

//                rowAll = resAll->querySubObject("Choose()");
//                rowAll->dynamicCall("Next()").toBool();
//                kodKarty = rowAll->dynamicCall("KodKarty").toString().simplified();

//                queryAll->clear();
//                delete queryAll;

//                queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
//                queryAll->dynamicCall("Text", queryTextDolzn.arg(kodSsylka));
//                resAll = queryAll->querySubObject("Execute()");
//                rowAll = resAll->querySubObject("Choose()");

                QSqlQuery queryEmp;
                orgId = "OWN000000001";

//                rowAll->dynamicCall("Next()").toBool();
//                subName = rowAll->dynamicCall("Subdivision").toString().simplified();
//                postName = rowAll->dynamicCall("Post").toString().simplified();
//                tabNumber = rowAll->dynamicCall("TabelNumber").toString();
//                datePostupleniya = rowAll->dynamicCall("DatePostupleniya").toDate();
//                dateUvolneniya = rowAll->dynamicCall("DateUvolneniya").toDate();
//                grafikRabot = rowAll->dynamicCall("GrafikRabot").toString().simplified();

                queryEmp.prepare("SELECT * FROM employee WHERE (employeename = :nameperson AND datebirthday = :datebirth);");
                queryEmp.bindValue(":nameperson",personName);
                queryEmp.bindValue(":datebirth",dateName);
                queryEmp.exec();
                queryEmp.next();

                /*if(queryPer.value(3).toString() != subId || queryPer.value(4).toString() != postId || queryPer.value(6).toString() != kodKarty){*/
                /*QSqlQuery lastQuery;
                        lastQuery.prepare("SELECT "
                                          "(SELECT subdivisionid FROM subdivision WHERE subdivisionname = :subname), "
                                          "(SELECT postid FROM post WHERE postname = :postname), "
                                          "kodkarty FROM employee WHERE employeeid = :employeeid;");
                        lastQuery.bindValue(":subname",subName);
                        lastQuery.bindValue(":postname",postName);
                        lastQuery.bindValue(":kodkarty",kodKarty);
                        lastQuery.bindValue(":employeeid",queryEmp.value(0).toString());
                        lastQuery.exec();
                        lastQuery.next();
                        if(lastQuery.isValid()){
                            if(lastQuery.value(0).toString() == subId || lastQuery.value(1).toString() == postId || lastQuery.value(2).toString() == kodKarty){
                                ++lll;
                                continue;
                            }
                        }*/
                //�������������
                if(subName != ""){
                    QSqlQuery querySub;
                    querySub.prepare("SELECT subdivisionid FROM subdivision WHERE subdivisionname = :subname AND organizationid = :orgid;");
                    querySub.bindValue(":subname",subName);
                    querySub.bindValue(":orgid",orgId);
                    querySub.exec();
                    querySub.next();
                    if(!querySub.isValid()){
                        QSqlQuery query;
                        NumPrefix numPrefix(this);
                        QString idSub = numPrefix.getPrefix("subdivision");
                        query.prepare("INSERT INTO subdivision (subdivisionid, subdivisionname, organizationid) VALUES(:idsub, :namesub, :orgid);");
                        query.bindValue(":namesub",subName);
                        query.bindValue(":orgid",orgId);
                        query.bindValue(":idsub",idSub);
                        query.exec();
                        subId = idSub;
                    }else{
                        subId = querySub.value(0).toString();
                    }
                }else{
                    subId = "OWN000000001";
                }

                if(postName != ""){
                    QSqlQuery queryPost;
                    queryPost.prepare("SELECT postid FROM post WHERE postname = :postname;");
                    queryPost.bindValue(":postname",postName);
                    queryPost.exec();
                    queryPost.next();
                    if(!queryPost.isValid()){
                        NumPrefix numPrefix(this);
                        QString idPost = numPrefix.getPrefix("post");
                        QSqlQuery query;
                        query.prepare("INSERT INTO post (postid, postname) VALUES(:idpost, :namepost);");
                        query.bindValue(":namepost",postName);
                        query.bindValue(":idpost", idPost);
                        query.exec();
                        postId = idPost;
                    }else{
                        postId = queryPost.value(0).toString();
                    }
                }else{
                    postId = "OWN000000001";
                }
                if(grafikRabot != ""){
                    QSqlQuery queryGR;
                    queryGR.prepare("SELECT grafikrabotid FROM grafikrabot WHERE grafikrabotname = :grafikname;");
                    queryGR.bindValue(":grafikname",grafikRabot);
                    queryGR.exec();
                    queryGR.next();
                    if(!queryGR.isValid()){
                        NumPrefix numPrefix(this);
                        QString idGR = numPrefix.getPrefix("grafikrabot");
                        QSqlQuery query;
                        query.prepare("INSERT INTO grafikrabot (grafikrabotid, grafikrabotname) VALUES(:id, :name);");
                        query.bindValue(":name",grafikRabot);
                        query.bindValue(":id", idGR);
                        query.exec();
                        grafikRabotId = idGR;
                    }else{
                        grafikRabotId = queryGR.value(0).toString();
                    }
                }else{
                    grafikRabotId = "OWN000000001";
                }
                QSqlQuery queryPer;
                queryPer.prepare("SELECT * FROM employee "
                                 "WHERE employeename = :personname AND "
                                 "organizationid = :orgid AND "
                                 "datebirthday = :datebirth;");
                queryPer.bindValue(":personname",personName);
                queryPer.bindValue(":orgid",orgId);
                queryPer.bindValue(":datebirth",dateName);
                queryPer.exec();
                queryPer.next();

                queryAll->clear();
                delete queryAll;

                queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
                queryAll->dynamicCall("Text", queryTextPassport.arg(kodSsylka));
                resAll = queryAll->querySubObject("Execute()");
                rowAll = resAll->querySubObject("Choose()");
                if(rowAll->dynamicCall("Next()").toBool()){
                    passportSeriya = rowAll->dynamicCall("PassportSeriya").toString().simplified();
                    passportNomer = rowAll->dynamicCall("PassportNomer").toString().simplified();
                    passportDataVidachi = rowAll->dynamicCall("PassportDataVidachi").toString().simplified();
                    kemVidan = rowAll->dynamicCall("KemVidan").toString().simplified();
                }
                if(!queryPer.isValid()){
                    QSqlQuery query;
                    NumPrefix numPrefix(this);
                    personId = numPrefix.getPrefix("employee");
                    query.prepare("INSERT INTO employee ("
                                  "employeeid, "
                                  "employeename, "
                                  "organizationid, "
                                  "subdivisionid, "
                                  "postid, "
                                  "tabnumber, "
                                  "kodkarty, "
                                  "datebirthday, "
                                  "grafikrabot, "
                                  "passportseriya, "
                                  "passportnumber, "
                                  "passportissueddate, "
                                  "withorganization, "
                                  "passportissuedby"
                                  ") VALUES("
                                  ":perid, "
                                  ":pername, "
                                  ":orgid, "
                                  ":subid, "
                                  ":postid, "
                                  ":tnumber, "
                                  ":kkarty, "
                                  ":birthday, "
                                  ":grafikrabot, "
                                  ":passportseriya, "
                                  ":passportnumber, "
                                  ":passportissueddate, "
                                  ":with, "
                                  ":issue"
                                  ");");
                    query.bindValue(":personid",personId);
                    query.bindValue(":pername",personName);
                    query.bindValue(":orgid",orgId);
                    query.bindValue(":subid",subId);
                    query.bindValue(":postid", postId);
                    query.bindValue(":tnumber", tabNumber);
                    query.bindValue(":kkarty",kodKarty);
                    query.bindValue(":birthday",dateName);
                    query.bindValue(":grafikrabot",grafikRabotId);
                    query.bindValue(":passportseriya",passportSeriya);
                    query.bindValue(":passportnumber",passportNomer);
                    query.bindValue(":passportissueddate",passportDataVidachi);
                    query.bindValue(":with",datePostupleniya);
                    query.bindValue(":issue",kemVidan);
                    query.exec();
                    if(!query.isActive()){
                        QMessageBox::warning(this,QObject::tr("DataBase ERROR!"),query.lastError().text());
                    }
                }else{
                    /*if(queryPer.value(3).toString() != subId || queryPer.value(4).toString() != postId || queryPer.value(6).toString() != kodKarty){*/
                    personId = queryPer.value(0).toString();

                    QSqlQuery query6;
                    query6.prepare("UPDATE employee SET "
                                   "kodkarty = :kodkart, "
                                   "subdivisionid = :subid, "
                                   "postid = :postid, "
                                   "grafikrabot = :grafikrabot, "
                                   "passportseriya = :passportseriya, "
                                   "passportnumber = :passportnumber, "
                                   "passportissueddate = :passportissueddate, "
                                   "withorganization = :with, "
                                   "passportissuedby = :issue "
                                   "WHERE (employeeid = :personid);");
                    query6.bindValue(":personid",personId);
                    query6.bindValue(":kodkart",kodKarty);
                    query6.bindValue(":subid",subId);
                    query6.bindValue(":postid", postId);
                    query6.bindValue(":grafikrabot",grafikRabotId);
                    query6.bindValue(":passportseriya",passportSeriya);
                    query6.bindValue(":passportnumber",passportNomer);
                    query6.bindValue(":passportissueddate",passportDataVidachi);
                    query6.bindValue(":with",datePostupleniya);
                    query6.bindValue(":issue",kemVidan);
                    query6.exec();
                    if(!query6.isActive()){
                        QMessageBox::warning(this,QObject::tr("DataBase ERROR!"),query6.lastError().text());
                    }
                }

                queryAll->clear();
                delete queryAll;

                queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
                queryAll->dynamicCall("Text", queryTextObrazovanie.arg(kodSsylka));
                resAll = queryAll->querySubObject("Execute()");
                rowAll = resAll->querySubObject("Choose()");
                while(rowAll->dynamicCall("Next()").toBool()){
                    vidObrazovaniya = rowAll->dynamicCall("VidObrazovaniya").toString().simplified();
                    uchebnoeZavedenie = rowAll->dynamicCall("UchebnoeZavedenie").toString().simplified();
                    spezialnost = rowAll->dynamicCall("Spezialnost").toString().simplified();
                    diplom = rowAll->dynamicCall("Diplom").toString().simplified();
                    godOconchaniya = convertStToDate(rowAll->dynamicCall("GodOconchaniya").toString().simplified());
                    kvalificaziya = rowAll->dynamicCall("Kvalificaziya").toString().simplified();
                    QSqlQuery queryOrg;
                    queryOrg.prepare("SELECT * FROM education WHERE (employeeid = :id AND typeofeducation = :type);");
                    queryOrg.bindValue(":id",personId);
                    queryOrg.bindValue(":type",vidObrazovaniya);
                    queryOrg.exec();
                    queryOrg.next();
                    if(!queryOrg.isValid()){
                        QSqlQuery query;
                        NumPrefix numPrefix(this);
                        QString idEdu = numPrefix.getPrefix("education");
                        query.prepare("INSERT INTO education ("
                                      "educationid, "
                                      "employeeid, "
                                      "typeofeducation, "
                                      "educationalinstitution, "
                                      "specialty, "
                                      "diplom, "
                                      "year, "
                                      "qualification"
                                      ") VALUES(:edid, :empid, :typeofeducation,"
                                      ":educationalinstitution, :specialty, :diplom, :year, :qualification);");
                        query.bindValue(":edid",idEdu);
                        query.bindValue(":empid",personId);
                        query.bindValue(":typeofeducation",vidObrazovaniya);
                        query.bindValue(":educationalinstitution",uchebnoeZavedenie);
                        query.bindValue(":specialty",spezialnost);
                        query.bindValue(":diplom",diplom);
                        query.bindValue(":year",godOconchaniya);
                        query.bindValue(":qualification",kvalificaziya);
                        query.exec();
                        if(!query.isActive()){
                            QMessageBox::warning(this,QObject::tr("Education ERROR!"),query.lastError().text());
                        }
                    }
                }

                queryAll->clear();
                delete queryAll;

                queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
                queryAll->dynamicCall("Text", queryTextPhone.arg(kodSsylka));
                resAll = queryAll->querySubObject("Execute()");
                rowAll = resAll->querySubObject("Choose()");
                while(rowAll->dynamicCall("Next()").toBool()){
                    tephone = rowAll->dynamicCall("Tephone").toString().simplified();
                    QSqlQuery queryOrg;
                    queryOrg.prepare("SELECT * FROM personalinformation WHERE (employeeid = :id AND information = :inf);");
                    queryOrg.bindValue(":id",personId);
                    queryOrg.bindValue(":inf",tephone);
                    queryOrg.exec();
                    queryOrg.next();
                    if(!queryOrg.isValid()){
                        QSqlQuery query;
                        NumPrefix numPrefix(this);
                        QString idTe = numPrefix.getPrefix("personalinformation");
                        query.prepare("INSERT INTO personalinformation ("
                                      "personalinformationid, "
                                      "employeeid, "
                                      "information"
                                      ") VALUES(:perid, :empid, :information);");
                        query.bindValue(":edid",idTe);
                        query.bindValue(":empid",personId);
                        query.bindValue(":information",tephone);
                        query.exec();
                        if(!query.isActive()){
                            QMessageBox::warning(this,QObject::tr("DataBase ERROR!"),query.lastError().text());
                        }
                    }
                }

                queryAll->clear();
                delete queryAll;

                queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
                queryAll->dynamicCall("Text", queryTextRodstva.arg(kodSsylka));
                resAll = queryAll->querySubObject("Execute()");
                rowAll = resAll->querySubObject("Choose()");
                while(rowAll->dynamicCall("Next()").toBool()){
                    stepenRodstva = rowAll->dynamicCall("StepenRodstva").toString().simplified();
                    nameRodstva = rowAll->dynamicCall("NameRodstva").toString().simplified();
                    dateRodstva = convertStToDate(rowAll->dynamicCall("DateRodstva").toString().simplified());
                    QSqlQuery queryOrg;
                    queryOrg.prepare("SELECT * FROM rodnie WHERE (employeeid = :id AND stepenrodstva = :stepen "
                                     "AND namerodstva = :name);");
                    queryOrg.bindValue(":id",personId);
                    queryOrg.bindValue(":stepen",stepenRodstva);
                    queryOrg.bindValue(":name",nameRodstva);
                    queryOrg.exec();
                    queryOrg.next();
                    if(!queryOrg.isValid()){
                        QSqlQuery query;
                        NumPrefix numPrefix(this);
                        QString idEdu = numPrefix.getPrefix("rodnie");
                        query.prepare("INSERT INTO rodnie ("
                                      "rodnieid, "
                                      "employeeid, "
                                      "stepenrodstva, "
                                      "namerodstva, "
                                      "birthrodstva"
                                      ") VALUES(:edid, :empid, :stepen,"
                                      ":name, :birth);");
                        query.bindValue(":edid",idEdu);
                        query.bindValue(":empid",personId);
                        query.bindValue(":stepen",stepenRodstva);
                        query.bindValue(":name",nameRodstva);
                        query.bindValue(":birth",dateRodstva);
                        query.exec();
                        if(!query.isActive()){
                            QMessageBox::warning(this,QObject::tr("Rodnie!"),query.lastError().text());
                        }
                    }

                }
                queryAll->clear();
                delete queryAll;
            }
            ++lll;
            qApp->processEvents();
            progressR.show();
        }
        excel->clear();
    }else{
        QMessageBox::warning(this,QObject::tr("Attention!!!"),QObject::tr("Don't create COM connector..."));
    }
    direc.removeDateBase();
}

QString ImportOneForm::convertStToDate(QString godO)
{
    QDate godOcon = QDate::fromString(godO,"dd.MM.yyyy");
    QString godOconchaniya;
    if(!godOcon.isValid()){
        godOcon = QDate::fromString(godO,"MM.yyyy");
    }
    if(!godOcon.isValid()){
        godOcon = QDate::fromString(godO,"yyyy");
    }
    if(!godOcon.isValid()){
        godOconchaniya = "100-01-01";
    }else{
        godOconchaniya = godOcon.toString(Qt::ISODate);
    }
    return godOconchaniya;
}
