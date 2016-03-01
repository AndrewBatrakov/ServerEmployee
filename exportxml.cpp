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

//        queryTextFizLiz = tr("ВЫБРАТЬ "
//                             "ФизическиеЛица.ПометкаУдаления КАК DeletionMark, "
//                             "ФизическиеЛица.Код КАК Kod, "
//                             "ФизическиеЛица.Наименование КАК FIO, "
//                             "ФизическиеЛица.ДатаРождения КАК DataRozdeniya "
//                             "ИЗ Справочник.ФизическиеЛица КАК ФизическиеЛица "
//                             "ГДЕ "
//                             "ФизическиеЛица.ЭтоГруппа <> ИСТИНА "
//                             "УПОРЯДОЧИТЬ ПО "
//                             "FIO "
//                             "АВТОУПОРЯДОЧИВАНИЕ");
        QByteArray text = ("ВЫБРАТЬ "
                             "ФизическиеЛица.Код КАК Kod, "
                             "ФизическиеЛица.Наименование КАК FIO, "
                             "ФизическиеЛица.ДатаРождения КАК DataRozdeniya, "
                             "ФизическиеЛица.ПометкаУдаления КАК DeletionMark, "
                             "РаботникиОрганизацийСрезПоследних.Сотрудник.Код КАК TabelNumber, "
                             "РаботникиОрганизацийСрезПоследних.ОбособленноеПодразделение.Наименование КАК Obosobl, "
                             "РаботникиОрганизацийСрезПоследних.ПодразделениеОрганизации.Наименование КАК Subdivision, "
                             "РаботникиОрганизацийСрезПоследних.Должность.Наименование КАК Post, "
                             "РаботникиОрганизацийСрезПоследних.ГрафикРаботы.Наименование КАК GrafikRabot, "
                             "РаботникиОрганизацийСрезПоследних.Сотрудник.ДатаПриемаНаРаботу КАК DatePostupleniya, "
                             "РаботникиОрганизацийСрезПоследних.Сотрудник.ДатаУвольнения КАК DateUvolneniya, "
                             "СИТ_АРМ_ДанныеКартСрезПоследних.КодКарты КАК KodKarty "
                             "ИЗ "
                             "Справочник.ФизическиеЛица КАК ФизическиеЛица "
                             "ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.РаботникиОрганизаций.СрезПоследних КАК РаботникиОрганизацийСрезПоследних "
                             "ПО ФизическиеЛица.Ссылка = РаботникиОрганизацийСрезПоследних.Сотрудник.Физлицо.Ссылка "
                             "ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.СИТ_АРМ_ДанныеКарт.СрезПоследних КАК СИТ_АРМ_ДанныеКартСрезПоследних "
                             "ПО ФизическиеЛица.Ссылка = СИТ_АРМ_ДанныеКартСрезПоследних.Физлицо.Ссылка "
                             "ГДЕ "
                             "НЕ ФизическиеЛица.ЭтоГруппа "
                             "И НЕ ФизическиеЛица.ПометкаУдаления "
                             "И РаботникиОрганизацийСрезПоследних.Сотрудник.ДатаУвольнения = ДАТАВРЕМЯ(1, 1, 1) "
                             "И НЕ РаботникиОрганизацийСрезПоследних.Сотрудник.ДатаПриемаНаРаботу = ДАТАВРЕМЯ(1, 1, 1) "
                             "И СИТ_АРМ_ДанныеКартСрезПоследних.Активна "
                             "УПОРЯДОЧИТЬ ПО "
                             "FIO");
        queryTextFizLiz = codecTr->toUnicode(text);

//        text = tr("ВЫБРАТЬ ПЕРВЫЕ 1 "
//                            "СИТ_АРМ_ДанныеКартСрезПоследних.КодКарты КАК KodKarty "
//                            "ИЗ "
//                            "РегистрСведений.СИТ_АРМ_ДанныеКарт.СрезПоследних КАК "
//                            "СИТ_АРМ_ДанныеКартСрезПоследних "
//                            "ГДЕ СИТ_АРМ_ДанныеКартСрезПоследних.Физлицо.Код = \"%1\" "
//                            "И СИТ_АРМ_ДанныеКартСрезПоследних.Активна = ИСТИНА "
//                            "УПОРЯДОЧИТЬ ПО "
//                            "СИТ_АРМ_ДанныеКартСрезПоследних.Период УБЫВ ");
//        queryTextKarty = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "ПаспротныеДанныеФизЛицСрезПоследних.ДокументСерия КАК PassportSeriya, "
                "ПаспротныеДанныеФизЛицСрезПоследних.ДокументНомер КАК PassportNomer, "
                "ПаспротныеДанныеФизЛицСрезПоследних.ДокументДатаВыдачи КАК PassportDataVidachi, "
                "ПаспротныеДанныеФизЛицСрезПоследних.ДокументКемВыдан КАК KemVidan "
                "ИЗ "
                "РегистрСведений.ПаспортныеДанныеФизЛиц.СрезПоследних КАК ПаспротныеДанныеФизЛицСрезПоследних "
                "ГДЕ "
                "ПаспротныеДанныеФизЛицСрезПоследних.ФизЛицо.Код = \"%1\"");
        queryTextPassport = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "РаботникиОрганизацийСрезПоследних.ПодразделениеОрганизации.Наименование КАК Subdivision, "
                "РаботникиОрганизацийСрезПоследних.Должность.Наименование КАК Post, "
                "РаботникиОрганизацийСрезПоследних.Сотрудник.Code КАК TabelNumber, "
                "РаботникиОрганизацийСрезПоследних.Сотрудник.ДатаПриемаНаРаботу КАК DatePostupleniya, "
                "РаботникиОрганизацийСрезПоследних.Сотрудник.ДатаУвольнения КАК DateUvolneniya, "
                "РаботникиОрганизацийСрезПоследних.Сотрудник.ГрафикРаботы.Наименование КАК GrafikRabot, "
                "РаботникиОрганизацийСрезПоследних.ОбособленноеПодразделение.Наименование КАК Obosobl "
                "ИЗ "
                "РегистрСведений.РаботникиОрганизаций.СрезПоследних КАК РаботникиОрганизацийСрезПоследних "
                "ГДЕ РаботникиОрганизацийСрезПоследних.Сотрудник.Физлицо.Код = \"%1\"");
        queryTextDolzn = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "ФизическиеЛицаОбразование.ВидОбразования.Наименование КАК VidObrazovaniya, "
                "ФизическиеЛицаОбразование.УчебноеЗаведение.Наименование КАК UchebnoeZavedenie, "
                "ФизическиеЛицаОбразование.Специальность.Наименование КАК Spezialnost, "
                "ФизическиеЛицаОбразование.Диплом КАК Diplom, "
                "ФизическиеЛицаОбразование.ГодОкончания КАК GodOconchaniya, "
                "ФизическиеЛицаОбразование.Квалификация КАК Kvalificaziya "
                "ИЗ "
                "Справочник.ФизическиеЛица.Образование КАК ФизическиеЛицаОбразование "
                "ГДЕ "
                "ФизическиеЛицаОбразование.Ссылка.Код = \"%1\"");
        queryTextObrazovanie = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "ФизическиеЛицаСоставСемьи.СтепеньРодства.Наименование КАК StepenRodstva, "
                "ФизическиеЛицаСоставСемьи.Имя КАК NameRodstva, "
                "ФизическиеЛицаСоставСемьи.ДатаРождения КАК BirthRodstva "
                "ИЗ "
                "Справочник.ФизическиеЛица.СоставСемьи КАК ФизическиеЛицаСоставСемьи "
                "ГДЕ "
                "ФизическиеЛицаСоставСемьи.Ссылка.Код = \"%1\"");
        queryTextRodstva = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "КонтактнаяИнформация.Представление КАК Tephone "
                "ИЗ "
                "РегистрСведений.КонтактнаяИнформация КАК КонтактнаяИнформация "
                "ГДЕ "
                "КонтактнаяИнформация.Объект.Код = \"%1\"");
        queryTextPhone = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ ПЕРВЫЕ 1 "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.НомерСправки КАК NomerSpravki, "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.Годен КАК Goden, "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.ДатаПрохождения КАК DataProhojdeniya, "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.Номер КАК NomerDocumenta, "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.МедосмотрПройден КАК Proyden, "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.Дата КАК Data "
                "ИЗ "
                "Документ.СИТ_АРМ_НаправленияИСведенияОМедОсмотре КАК СИТ_АРМ_НаправленияИСведенияОМедОсмотре "
                "ГДЕ "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.Сотрудник.Код = \"%1\" "
                "УПОРЯДОЧИТЬ ПО "
                "СИТ_АРМ_НаправленияИСведенияОМедОсмотре.Дата УБЫВ");
        queryTextMedosmotr = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ ПЕРВЫЕ 1 "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасностиПрограммыОбучения.ПрограммаОбучения.Наименование КАК ProgrammaObucheniya, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности.Номер КАК Nomer, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности.Дата КАК Data, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности.ОбъемЧасов КАК ObChasov "
                "ИЗ "
                "Документ.СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности.ПрограммыОбучения КАК СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасностиПрограммыОбучения "
                "ЛЕВОЕ СОЕДИНЕНИЕ Документ.СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности КАК СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности "
                "ПО СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасностиПрограммыОбучения.Ссылка = СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности.Ссылка "
                "ГДЕ "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПромышленнойБезопасности.Работники.Сотрудник.Физлицо.Код = \"%1\"");
        queryTextPromBezop = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ ПЕРВЫЕ 1 "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Дата КАК DataUdostovereniya, "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Протокол.Дата КАК ProtokolData, "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Протокол.ПрограммаОбучения.Наименование КАК ProtocolProgrammaObucheniya, "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Протокол.Номер КАК ProtocolNomer, "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Протокол.ОбъемЧасов КАК ProtokolObChasov "
                "ИЗ "
                "Документ.СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда КАК СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда "
                "ГДЕ "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Сотрудник.Код = \"%1\" "
                "УПОРЯДОЧИТЬ ПО "
                "СИТ_АРМ_УдостоверениеОПроверкеЗнанийТребованийОхраныТруда.Дата УБЫВ");
        queryTextOhranaTruda = codecTr->toUnicode(text);


        text = ("ВЫБРАТЬ ПЕРВЫЕ 1 "
                "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Дата КАК Data, "
                "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Номер КАК Number, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Номер КАК PTMNumber, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Дата КАК PTMDate, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.ОбъемЧасов КАК PTMObChasov, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.КраткийСоставДокумента КАК PTMSostav, "
                "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.ПрограммаОбучения.Наименование КАК PTMProgramma "
                "ИЗ "
                "Документ.СИТ_АРМ_ТалонПожарноТехническогоМинимума КАК СИТ_АРМ_ТалонПожарноТехническогоМинимума "
                "ЛЕВОЕ СОЕДИНЕНИЕ Документ.СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума КАК СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума "
                "ПО СИТ_АРМ_ТалонПожарноТехническогоМинимума.Протокол = СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Ссылка "
                "ГДЕ "
                "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Сотрудник.Код = \"%1\" "
                "УПОРЯДОЧИТЬ ПО "
                "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Дата УБЫВ");
        queryTextPTM = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.Рост КАК Rost, "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.ОдеждаЗимняя КАК OdezhdaZim, "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.ОдеждаЛетняя КАК OdezhdaLet, "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.ОбувьЗимняя КАК ObuvZim, "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.ОбувьЛетняя КАК ObuvLet, "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.ГоловнойУбор КАК GolUbor "
                "ИЗ "
                "Справочник.СИТ_АРМ_ЛичнаяКарточкаСотрудника КАК СИТ_АРМ_ЛичнаяКарточкаСотрудника "
                "ГДЕ "
                "СИТ_АРМ_ЛичнаяКарточкаСотрудника.Владелец.Физлицо.Код = \"%1\"");
        queryTextRost = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "ФизическиеЛицаСтажи.ВидСтажа.Наименование, "
                "ФизическиеЛицаСтажи.ДатаОтсчета, "
                "ФизическиеЛицаСтажи.РазмерМесяцев, "
                "ФизическиеЛицаСтажи.РазмерДней "
                "ИЗ "
                "Справочник.ФизическиеЛица.Стажи КАК ФизическиеЛицаСтажи "
                "ГДЕ "
                "ФизическиеЛицаСтажи.Ссылка.Код = \"%1\"");
        queryTextStaz = codecTr->toUnicode(text);

        text = ("ВЫБРАТЬ "
                "ФизическиеЛицаПрофессии.Профессия "
                "ИЗ "
                "Справочник.ФизическиеЛица.Профессии КАК ФизическиеЛицаПрофессии "
                "ГДЕ "
                "ФизическиеЛицаПрофессии.Ссылка.Код = \"%1\"");
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
        //xml.writeAttribute(tr("Всего"),QString::number(cou));
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
