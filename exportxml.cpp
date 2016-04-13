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
            "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Номер КАК ptmUdNumber, "
            "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Дата КАК ptmUdData, "
            "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Номер КАК ptmNumber, "
            "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Дата КАК ptmDate, "
            "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.ОбъемЧасов КАК ptmObChasov, "
            "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.ПрограммаОбучения.Наименование КАК ptmProgramma, "
            "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.КраткийСоставДокумента КАК ptmSostav "
            "ИЗ "
            "Документ.СИТ_АРМ_ТалонПожарноТехническогоМинимума КАК СИТ_АРМ_ТалонПожарноТехническогоМинимума "
            "ЛЕВОЕ СОЕДИНЕНИЕ Документ.СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума КАК СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума "
            "ПО СИТ_АРМ_ТалонПожарноТехническогоМинимума.Протокол = СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Ссылка "
            "ГДЕ "
            "СИТ_АРМ_ТалонПожарноТехническогоМинимума.Сотрудник.Код = \"%1\" "
            "УПОРЯДОЧИТЬ ПО "
            "СИТ_АРМ_ПротоколПроверкиЗнанийПожарноТехническогоМинимума.Дата "
            "ptmNumber УБЫВ");

    queryTextPTM = codecTr->toUnicode(text);

    text = ("ВЫБРАТЬ ПЕРВЫЕ 1 "
            "СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте.Номер КАК NumberUd, "
            "СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте.Дата КАК DataUd, "
            "СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте.Протокол.Номер КАК visNumber, "
            "СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте.Протокол.Дата КАК visData, "
            "СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте.Протокол.ОбъемЧасов КАК visObChasov "
            "ИЗ "
            "Документ.СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте КАК СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте "
            "ГДЕ "
            "СИТ_АРМ_УдостоверениеОбучененияМетодамРаботыНаВысоте.Сотрудник.Код = \"%1\" "
            "УПОРЯДОЧИТЬ ПО "
            "DataUd УБЫВ");
    queryTextVis = codecTr->toUnicode(text);

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
            "СИТ_АРМ_МатериалыВЭксплуатацииОстатки.СредствоИндивидуальнойЗащиты.Наименование КАК Naim, "
            "СИТ_АРМ_МатериалыВЭксплуатацииОстатки.КоличествоОстаток КАК Ostatok, "
            "СИТ_АРМ_МатериалыВЭксплуатацииОстатки.ДатаПередачи КАК DataPeredachi, "
            "СИТ_АРМ_МатериалыВЭксплуатацииОстатки.СрокИспользования КАК SrokIspolzovaniya "
            "ИЗ "
            "РегистрНакопления.СИТ_АРМ_МатериалыВЭксплуатации.Остатки КАК СИТ_АРМ_МатериалыВЭксплуатацииОстатки "
            "ГДЕ "
            "СИТ_АРМ_МатериалыВЭксплуатацииОстатки.Сотрудник.Код = \"%1\" ");

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
