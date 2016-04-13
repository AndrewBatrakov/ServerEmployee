#include "exportxml.h"
#include <qaxobject.h>
#include "putftp.h"

ExportXML::ExportXML()
{
}

ExportXML::~ExportXML()
{
}

void ExportXML::openImport(bool allData)
{
    excel = new QAxObject("v83.ComConnector",0);
    if(excel){
        return;
    }
    queryAll = new QAxObject;
    resAll = new QAxObject;
    rowAll = new QAxObject;

    //QTextCodec::setCodecForCStrings(QTextCodec::codecForName("UTF8"));
    //QTextCodec::setCodecForTr(QTextCodec::codecForName("Windows-1251"));
    QTextCodec *codecTr = QTextCodec::codecForName("Windows-1251");
    QString queryTextFizLiz, queryTextVis, queryTextDolzn, queryTextObrazovanie, queryTextRodstva,
            queryTextPhone, queryTextMedosmotr, queryTextPromBezop, queryTextOhranaTruda, queryTextPTM, queryTextRost,
            queryTextPassport, queryOstatkiSIZ;


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
            "���_���_��������������������������������.����� ��� ptmUdNumber, "
            "���_���_��������������������������������.���� ��� ptmUdData, "
            "���_���_�������������������������������������������������.����� ��� ptmNumber, "
            "���_���_�������������������������������������������������.���� ��� ptmDate, "
            "���_���_�������������������������������������������������.���������� ��� ptmObChasov, "
            "���_���_�������������������������������������������������.�����������������.������������ ��� ptmProgramma, "
            "���_���_�������������������������������������������������.���������������������� ��� ptmSostav "
            "�� "
            "��������.���_���_�������������������������������� ��� ���_���_�������������������������������� "
            "����� ���������� ��������.���_���_������������������������������������������������� ��� ���_���_������������������������������������������������� "
            "�� ���_���_��������������������������������.�������� = ���_���_�������������������������������������������������.������ "
            "��� "
            "���_���_��������������������������������.���������.��� = \"%1\" "
            "����������� �� "
            "���_���_�������������������������������������������������.���� "
            "ptmNumber ����");

    queryTextPTM = codecTr->toUnicode(text);

    text = ("������� ������ 1 "
            "���_���_��������������������������������������������.����� ��� NumberUd, "
            "���_���_��������������������������������������������.���� ��� DataUd, "
            "���_���_��������������������������������������������.��������.����� ��� visNumber, "
            "���_���_��������������������������������������������.��������.���� ��� visData, "
            "���_���_��������������������������������������������.��������.���������� ��� visObChasov "
            "�� "
            "��������.���_���_�������������������������������������������� ��� ���_���_�������������������������������������������� "
            "��� "
            "���_���_��������������������������������������������.���������.��� = \"%1\" "
            "����������� �� "
            "DataUd ����");
    queryTextVis = codecTr->toUnicode(text);

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
            "���_���_�����������������������������.����������������������������.������������ ��� Naim, "
            "���_���_�����������������������������.����������������� ��� Ostatok, "
            "���_���_�����������������������������.������������ ��� DataPeredachi, "
            "���_���_�����������������������������.����������������� ��� SrokIspolzovaniya "
            "�� "
            "�����������������.���_���_����������������������.������� ��� ���_���_����������������������������� "
            "��� "
            "���_���_�����������������������������.���������.��� = \"%1\" ");

    queryOstatkiSIZ = codecTr->toUnicode(text);

    QAxObject *rrr = excel->querySubObject("Connect(QVariant &)", QVariant("Srvr=\"rif-1c\";Ref=\"1c_zupkorp\";Usr=\"ManilchukNM\";Pwd=\"1Q34WER\""));
    if(rrr){
        return;
    }
    QAxObject *queryFizLiz;
    QAxObject *resFizLiz;
    QAxObject *rowFizLiz;

    queryFizLiz = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
    if(queryFizLiz){
        return;
    }
    queryFizLiz->dynamicCall("Text", queryTextFizLiz);
    resFizLiz = queryFizLiz->querySubObject("Execute()");
    if(resFizLiz){
        return;
    }
    rowFizLiz = resFizLiz->querySubObject("Choose()");
    if(rowFizLiz){
        return;
    }

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
    file.open(QIODevice::WriteOnly);

    xml.setDevice(&file);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement(tr("FileExchange"));
    xml.writeAttribute(tr("Date"),QDateTime::currentDateTime().toString("dd.MM.yyyy HH:mm:ss"));
    xml.writeStartElement(tr("Employees"));

    QString subName, postName, tabNumber, datePostupleniya, dateUvolneniya, grafikRabot, obosobl;
    QString personName, dateName, kodSsylka, kodKarty, tephone;
    QString vidObrazovaniya, uchebnoeZavedenie, spezialnost, diplom, godOconchaniya, kvalificaziya;
    QString stepenRodstva,nameRodstva,dateRodstva;
    QString sizName, sizOstatok, sizData, DataPeredachi, sizSrok;

    while(rowFizLiz->dynamicCall("Next()").toBool()){
        if(rowFizLiz->dynamicCall("DeletionMark").toBool()){
            continue;
        }

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
        xml.writeAttribute(tr("GrafikRabot"),grafikRabot);

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

            int number = 1;
            QString tempString;
            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextObrazovanie.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");

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
            queryAll->dynamicCall("Text", queryTextMedosmotr.arg(tabNumber));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
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
            queryAll->dynamicCall("Text", queryOstatkiSIZ.arg(tabNumber));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            number = 1;
            while(rowAll->dynamicCall("Next()").toBool()){
                sizName  = rowAll->dynamicCall("Naim").toString().simplified();
                sizOstatok = rowAll->dynamicCall("Ostatok").toString().simplified();
                sizData = rowAll->dynamicCall("DataPeredachi").toString();
                sizSrok = rowAll->dynamicCall("SrokIspolzovaniya").toString();

                QDateTime timeTemp = QDateTime::fromString(sizData,"yyyy-MM-ddTHH:mm:ss");
                QDateTime tt = timeTemp.addMonths(sizSrok.toInt());
                if(tt < QDateTime::currentDateTime()){
                    continue;
                }
                tempString = "sizName";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,sizName);

                tempString = "sizOstatok";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,sizOstatok);

                tempString = "sizData";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,sizData);

                tempString = "sizSrok";
                tempString += QString::number(number);
                xml.writeAttribute(tempString,sizSrok);
                ++number;
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextPromBezop.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            while(rowAll->dynamicCall("Next()").toBool()){
                xml.writeAttribute("pbProgrammaObucheniya",rowAll->dynamicCall("ProgrammaObucheniya").toString().simplified());
                xml.writeAttribute("pbNomer",rowAll->dynamicCall("Nomer").toString().simplified());
                xml.writeAttribute("pbData",rowAll->dynamicCall("Data").toString().simplified());
                xml.writeAttribute("pbObChasov",rowAll->dynamicCall("ObChasov").toString().simplified());

            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextOhranaTruda.arg(tabNumber));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            while(rowAll->dynamicCall("Next()").toBool()){
                xml.writeAttribute("otDataUdostovereniya",rowAll->dynamicCall("DataUdostovereniya").toString().simplified());
                xml.writeAttribute("otProtokolData",rowAll->dynamicCall("ProtokolData").toString().simplified());
                xml.writeAttribute("otProtocolProgrammaObucheniya",rowAll->dynamicCall("ProtocolProgrammaObucheniya").toString().simplified());
                xml.writeAttribute("otProtocolNomer",rowAll->dynamicCall("ProtocolNomer").toString().simplified());
                xml.writeAttribute("otProtokolObChasov",rowAll->dynamicCall("ProtokolObChasov").toString().simplified());
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextPTM.arg(tabNumber));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            while(rowAll->dynamicCall("Next()").toBool()){
                xml.writeAttribute("ptmUdData",rowAll->dynamicCall("ptmUdData").toString().simplified());
                xml.writeAttribute("ptmUdNumber",rowAll->dynamicCall("ptmUdNumber").toString().simplified());
                xml.writeAttribute("ptmNumber",rowAll->dynamicCall("ptmNumber").toString().simplified());
                xml.writeAttribute("ptmDate",rowAll->dynamicCall("ptmDate").toString().simplified());
                xml.writeAttribute("ptmObChasov",rowAll->dynamicCall("ptmObChasov").toString().simplified());
                xml.writeAttribute("ptmSostav",rowAll->dynamicCall("ptmSostav").toString().simplified());
                xml.writeAttribute("ptmProgramma",rowAll->dynamicCall("ptmProgramma").toString().simplified());
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextRost.arg(kodSsylka));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            while(rowAll->dynamicCall("Next()").toBool()){
                xml.writeAttribute("Rost",rowAll->dynamicCall("Rost").toString().simplified());
                xml.writeAttribute("OdezhdaZim",rowAll->dynamicCall("OdezhdaZim").toString().simplified());
                xml.writeAttribute("OdezhdaLet",rowAll->dynamicCall("OdezhdaLet").toString().simplified());
                xml.writeAttribute("ObuvZim",rowAll->dynamicCall("ObuvZim").toString().simplified());
                xml.writeAttribute("ObuvLet",rowAll->dynamicCall("ObuvLet").toString().simplified());
                xml.writeAttribute("GolUbor",rowAll->dynamicCall("GolUbor").toString().simplified());
            }
            queryAll->clear();
            delete queryAll;

            queryAll = rrr->querySubObject("NewObject(QVariant &)",QVariant(tr("Query")));
            queryAll->dynamicCall("Text", queryTextVis.arg(tabNumber));
            resAll = queryAll->querySubObject("Execute()");
            rowAll = resAll->querySubObject("Choose()");
            while(rowAll->dynamicCall("Next()").toBool()){
                xml.writeAttribute("numberUd",rowAll->dynamicCall("NumberUd").toString().simplified());
                xml.writeAttribute("dataUd",rowAll->dynamicCall("DataUd").toString().simplified());
                xml.writeAttribute("visNumber",rowAll->dynamicCall("visNumber").toString().simplified());
                xml.writeAttribute("visData",rowAll->dynamicCall("visData").toString().simplified());
                xml.writeAttribute("visObChasov",rowAll->dynamicCall("visObChasov").toString().simplified());
            }
            queryAll->clear();
            delete queryAll;

        }

        xml.writeEndElement();
    }

    xml.writeEndElement();
    xml.writeEndDocument();

    file.close();
    QFile fileOut;
    fileOut.setFileName(file.fileName());
    fileOut.open(QIODevice::ReadOnly);
    QByteArray byteArray = fileOut.readAll();
    if(byteArray.isEmpty()){
        qDebug() << fileOut.errorString();
    }


    QByteArray compressData = qCompress(byteArray);
    QFile compressFile;
    QFileInfo fileInfo;
    fileInfo.setFile(file.fileName());
    QString fN = fileInfo.baseName();
    fN += ".arh";
    compressFile.setFileName(fN);
    compressFile.open(QIODevice::WriteOnly);
    compressFile.write(compressData);

    fileOut.close();
    compressFile.close();

    excel->clear();
    excel->~QAxObject();

    PutFtp *putFtp = new PutFtp;
    qDebug()<<compressFile.fileName()<<", "<<QTime::currentTime();
    putFtp->putFile(compressFile.fileName());

    PutFtp *putFtp1 = new PutFtp;
    QString nullFileName = "Null";
    nullFileName += fN;
    QFile nullFile;
    nullFile.setFileName(nullFileName);
    nullFile.open(QIODevice::WriteOnly);
    QByteArray rr = "null";
    nullFile.write(rr);
    nullFile.close();
    putFtp1->putFile(nullFile.fileName());
    qDebug()<<nullFile.fileName()<<", "<<QTime::currentTime();
    compressFile.remove();
    fileOut.remove();
    nullFile.remove();

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
