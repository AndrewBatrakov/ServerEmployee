#include <QtWidgets>
#include <QtSql>

#include "exportxml.h"
#include "numprefix.h"
#include "mainwindow.h"
#include <qaxobject.h>
#include "databasedirection.h"
#include "putftp.h"

ExportXML::ExportXML(QObject *parent)
    :QObject(parent)
{
    excel = new QAxObject("v83.ComConnector",this);
    queryAll = new QAxObject;
    resAll = new QAxObject;
    rowAll = new QAxObject;
}

void ExportXML::openImport(bool allData)
{
    //QTextCodec::setCodecForCStrings(QTextCodec::codecForName("UTF8"));
    //QTextCodec::setCodecForTr(QTextCodec::codecForName("Windows-1251"));
    QTextCodec *codecTr = QTextCodec::codecForName("Windows-1251");

    if(!excel->isNull() == true){
        qApp->processEvents();
        QString queryTextFizLiz, queryTextKarty, queryTextDolzn, queryTextObrazovanie, queryTextRodstva, queryTextStaz,
                queryTextPhone, queryTextMedosmotr, queryTextPromBezop, queryTextOhranaTruda, queryTextPTM, queryTextRost,
                queryTextProfessii, queryTextPassport;

//        queryTextFizLiz = tr("������� "
//                             "��������������.��������������� ��� DeletionMark, "
//                             "��������������.��� ��� Kod, "
//                             "��������������.������������ ��� FIO, "
//                             "��������������.������������ ��� DataRozdeniya "
//                             "�� ����������.�������������� ��� �������������� "
//                             "��� "
//                             "��������������.��������� <> ������ "
//                             "����������� �� "
//                             "FIO "
//                             "������������������");
        QByteArray text = ("������� "
                             "��������������.��� ��� Kod, "
                             "��������������.������������ ��� FIO, "
                             "��������������.������������ ��� DataRozdeniya, "
                             "��������������.��������������� ��� DeletionMark, "
                             "���������������������������������.���������.��� ��� TabelNumber, "
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

//        text = tr("������� ������ 1 "
//                            "���_���_�����������������������.�������� ��� KodKarty "
//                            "�� "
//                            "���������������.���_���_����������.������������� ��� "
//                            "���_���_����������������������� "
//                            "��� ���_���_�����������������������.�������.��� = \"%1\" "
//                            "� ���_���_�����������������������.������� = ������ "
//                            "����������� �� "
//                            "���_���_�����������������������.������ ���� ");
//        queryTextKarty = codecTr->toUnicode(text);

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
                "���������������������������������.������������������������.������������ ��� Subdivision, "
                "���������������������������������.���������.������������ ��� Post, "
                "���������������������������������.���������.Code ��� TabelNumber, "
                "���������������������������������.���������.������������������ ��� DatePostupleniya, "
                "���������������������������������.���������.�������������� ��� DateUvolneniya, "
                "���������������������������������.���������.������������.������������ ��� GrafikRabot, "
                "���������������������������������.�������������������������.������������ ��� Obosobl "
                "�� "
                "���������������.��������������������.������������� ��� ��������������������������������� "
                "��� ���������������������������������.���������.�������.��� = \"%1\"");
        queryTextDolzn = codecTr->toUnicode(text);

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
        if(rrr){
        QAxObject *queryFizLiz = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
        queryFizLiz->dynamicCall("Text", queryTextFizLiz);
        QAxObject *resFizLiz = queryFizLiz->querySubObject("Execute()");
        QAxObject *rowFizLiz = resFizLiz->querySubObject("Choose()");

        //QProgressDialog progressR(this);

        //progressR.setWindowModality(Qt::WindowModal);
        //progressR.setMinimum(1);
        //int cou = rowFizLiz->dynamicCall("Count()").toInt();
        //progressR.setMaximum(cou);
        //progressR.activateWindow();

        QFile file;
        QXmlStreamWriter xml;
        if(allData){
            file.setFileName("ObmenAll.xml");
        }else{
            file.setFileName("Obmen.xml");
        }
        if(file.exists()){
            file.remove();
        }
        if(!file.open(QIODevice::WriteOnly)){
            //QMessageBox::warning(this,"Attention!","XML file do not open!!!");
            return;
        }

        xml.setDevice(&file);
        xml.setAutoFormatting(true);
        xml.writeStartDocument();
        xml.writeStartElement(tr("FileExchange"));
        xml.writeAttribute(tr("Date"),QDateTime::currentDateTime().toString("dd.MM.yyyy HH:mm:ss"));
        //xml.writeAttribute(tr("�����"),QString::number(cou));
        xml.writeStartElement(tr("Employees"));

        int lll = 0;
        QString subName, postName, tabNumber, datePostupleniya, dateUvolneniya, grafikRabot, obosobl;
        QString personName, dateName, kodSsylka, kodKarty, tephone;
        QString vidObrazovaniya, uchebnoeZavedenie, spezialnost, diplom, godOconchaniya, kvalificaziya;
        QString stepenRodstva,nameRodstva,dateRodstva;
        //bool onWork;


        while(rowFizLiz->dynamicCall("Next()").toBool()){
            //if(progressR.wasCanceled()){
            //    xml.writeEndElement();
            //    break;
            //}

            //progressR.setWindowTitle(QString::number(lll));
            //progressR.setValue(lll);
            //progressR.activateWindow();
            //progressR.show();

            if(rowFizLiz->dynamicCall("DeletionMark").toBool()){
                ++lll;
                continue;
            }

            qApp->processEvents();
            personName = rowFizLiz->dynamicCall("FIO").toString().simplified();
            dateName = rowFizLiz->dynamicCall("DataRozdeniya").toString();
            kodSsylka = rowFizLiz->dynamicCall("Kod").toString();
            subName = rowFizLiz->dynamicCall("Subdivision").toString().simplified();
            postName = rowFizLiz->dynamicCall("Post").toString().simplified();
            tabNumber = rowFizLiz->dynamicCall("TabelNumber").toString();
            datePostupleniya = rowFizLiz->dynamicCall("DatePostupleniya").toString();
            dateUvolneniya = rowFizLiz->dynamicCall("DateUvolneniya").toString();
            grafikRabot = rowFizLiz->dynamicCall("GrafikRabot").toString().simplified();
            obosobl = rowFizLiz->dynamicCall("Obosobl").toString().simplified();
            kodKarty = rowFizLiz->dynamicCall("KodKarty").toString().simplified();

            xml.writeStartElement(tr("Element"));
            xml.writeAttribute(tr("Employee"),personName);
            xml.writeAttribute(tr("DataRozdeniya"),dateName);
            xml.writeAttribute(tr("TabelNumber"),tabNumber);
            xml.writeAttribute(tr("Subdivision"),subName);
            xml.writeAttribute(tr("Post"),postName);
            xml.writeAttribute(tr("Obosobl"),obosobl);
            xml.writeAttribute(tr("DatePostupleniya"),datePostupleniya);
            xml.writeAttribute(tr("DateUvolneniya"),dateUvolneniya);
            xml.writeAttribute(tr("KodKarty"),kodKarty);

            if(allData){
            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextPassport.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");

            if(rowAll->dynamicCall("Next()").toBool()){
                xml.writeAttribute("PassportSeriya",rowAll->dynamicCall("PassportSeriya").toString().simplified());
                xml.writeAttribute("PassportNomer",rowAll->dynamicCall("PassportNomer").toString().simplified());
                xml.writeAttribute("PassportDataVidachi",rowAll->dynamicCall("PassportDataVidachi").toString());
                xml.writeAttribute("KemVidan",rowAll->dynamicCall("KemVidan").toString().simplified());
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextObrazovanie.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            int number = 1;
            QString tempString;
            while(rowAll->dynamicCall("Next()").toBool()){

                vidObrazovaniya = rowAll->dynamicCall("VidObrazovaniya").toString().simplified();
                uchebnoeZavedenie = rowAll->dynamicCall("UchebnoeZavedenie").toString().simplified();
                spezialnost = rowAll->dynamicCall("Spezialnost").toString().simplified();
                diplom = rowAll->dynamicCall("Diplom").toString().simplified();
                godOconchaniya = rowAll->dynamicCall("GodOconchaniya").toString();
                kvalificaziya = rowAll->dynamicCall("Kvalificaziya").toString().simplified();

                tempString = "VidObrazovaniya";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,vidObrazovaniya);

                tempString = "UchebnoeZavedenie";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,uchebnoeZavedenie);

                tempString = "Spezialnost";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,spezialnost);

                tempString = "Diplom";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,diplom);

                tempString = "GodOconchaniya";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,godOconchaniya);

                tempString = "Kvalificaziya";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,kvalificaziya);

                ++number;
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextPhone.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            number = 1;
            while(rowAll->dynamicCall("Next()").toBool()){
                tephone = rowAll->dynamicCall("Tephone").toString().simplified();

                tempString = "Tephone";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,tephone);
                ++number;
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextRodstva.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            number = 1;
            while(rowAll->dynamicCall("Next()").toBool()){

                stepenRodstva = rowAll->dynamicCall("StepenRodstva").toString().simplified();
                nameRodstva = rowAll->dynamicCall("NameRodstva").toString().simplified();
                dateRodstva = rowAll->dynamicCall("BirthRodstva").toString();

                tempString = "StepenRodstva";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,stepenRodstva);

                tempString = "NameRodstva";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,nameRodstva);

                tempString = "DateRodstva";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,dateRodstva);
                ++number;
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextMedosmotr.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            number = 1;
            while(rowAll->dynamicCall("Next()").toBool()){

                xml.writeAttribute("mNomerSpravki",rowAll->dynamicCall("NomerSpravki").toString().simplified());
                xml.writeAttribute("mGoden",rowAll->dynamicCall("Goden").toString().simplified());
                xml.writeAttribute("mDataProhojdeniya",rowAll->dynamicCall("DataProhojdeniya").toString());
                xml.writeAttribute("mNomerDocumenta",rowAll->dynamicCall("NomerDocumenta").toString());
                xml.writeAttribute("mProyden",rowAll->dynamicCall("Proyden").toString());
                xml.writeAttribute("mData",rowAll->dynamicCall("Data").toString());
            }
            queryAll->clear();
            delete queryAll;
            }

            xml.writeEndElement();
        }

        xml.writeEndElement();
        xml.writeEndDocument();

        file.close();
//        QFile fileOut("./Obmen.xml");
//        fileOut.open(QIODevice::ReadOnly);
//        QByteArray byteArray = fileOut.readAll();
//        if(byteArray.isEmpty()){
//            QMessageBox::warning(this,"Attention!!!",file.errorString());
//        }


        //QByteArray compressData = qCompress(byteArray);
        //QFile compressFile("./Obmen.rar");
//        compressFile.open(QIODevice::WriteOnly);
//        compressFile.write(compressData);

//        fileOut.close();
//        //file.remove();

//        compressFile.close();

        //QFile fileOut("./Obmen.rar");
   //fileOut.open(QIODevice::ReadOnly);
        //QByteArray bb = file.readAll();
        //QFile fO("./OO.xml");
        //fO.open(QIODevice::WriteOnly);
        //fO.write(qUncompress(bb));

        //fileOut.close();
        //fO.close();
        //while(compressFile->isOpen());

        excel->clear();
        excel->~QAxObject();

        PutFtp putFtp;
        qDebug()<<file.fileName();
        putFtp.putFile(file.fileName());
    }else{
        QTextStream ocout(stdout);
        ocout << "Don't create COM connector...";
    }
    }
}

QString ExportXML::convertStToDate(QString godO)
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
