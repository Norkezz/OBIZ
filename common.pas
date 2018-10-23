unit common;

interface
 uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,TLHelp32,ShellAPI,DateUtils,
  Dialogs, DB,  MyAccess, MemDS, Grids,  CRGrid, StdCtrls, Wininet,math,WinSock,IniFiles,GmXml,IdText,
  Xml.xmldom, Xml.XMLIntf, Xml.Win.msxmldom, Xml.XMLDoc,
  Mask, DBCtrls, Buttons,  MyScript,  Utilites, RpRave, uOpenOffice,IdMessage, IdSMTP,IdFTP,IdFTPCommon,
  RpDefine, RpCon, RpConDS, DAScript, DBAccess, DBGrids, RpBase, RpSystem,ADODB, IdAttachmentFile,IdHTTP,
  ExtCtrls, Menus, RpRender, RpRenderPDF, RpRenderRTF, ComCtrls, ComObj,main,RanderComLib_TLB;

const
  ConvertSet : array[0..255] of byte =
{таблица перекодировки ASCII с альтернативной кодовой страницой 866 в
WIN 1251. Украинские символы - по кодовой таблице PRINTFXU. Непечатные
символы заменяются пробелами}
{основная таблица}
{      00  01  02  03  04  05  06  07  08  09  0A  0B  0C  0D  0E  0F
{00} ( 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32,
{10}   32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32,
{20}   32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47,
{30}   48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63,
{40}   64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79,
{50}   80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95,
{60}   96, 97, 98, 99,100,101,102,103,104,105,106,107,108,109,110,111,
{70}  112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,
{дополнительная таблица}
{80}  192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,
{90}  208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,
{A0}  224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,
{B0}   32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32,
{C0}   32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32,
{B0}   32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32,
{E0}  240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255,
{F0}  168,184,178,179, 32, 32,175,191,170,186, 32,177,185, 32, 32, 32);


const
  NameMonth : array[0..11] of string =
  ('январь','февраль','март','апрель','май','июнь','июль','август',
    'сентябрь','октябрь','ноябрь','декабрь'
  );

const BodyScheet = '******   Система информационного оповещения ООО "Омскбланкиздат"' + #13 + #13 +
                'Настоящее письмо создано и отправлено роботом рассылки.' + #13 +
                'Не отвечайте на письмо, ЭТОТ АДРЕС НЕДОСТУПЕН ДЛЯ ВХОДЯЩЕЙ ПОЧТЫ.' + #13 +
                'E-mail для переписки уточняйте в отделе продаж по тел.: +7 3812 212-111,' + #13 +
                'или отправляйте сообщения через сайт http://www.omskblankizdat.ru' + #13 + #13 +
                '***** К письму прикреплен счет на оплату. *****' + #13 + #13 ;

const BodyAkt = '******   Система информационного оповещения ООО "Омскбланкиздат"' + #13 + #13 +
                'Настоящее письмо создано и отправлено роботом рассылки.' + #13 +
                'Не отвечайте на письмо, ЭТОТ АДРЕС НЕДОСТУПЕН ДЛЯ ВХОДЯЩЕЙ ПОЧТЫ.' + #13 +
                'E-mail для переписки уточняйте в бухгалтерии по тел.: +7 3812 240-677,' + #13 +
                'или прочитайте в тексте акта сверки.' + #13 + #13 +
                '***** К письму прикреплен акт сверки взаиморасчетов. *****' + #13 + #13 ;

const BodyCommon =
                'Внимание, открыть вложение возможно с помощью Adobe Acrobat Reader для чтения файлов в формате PDF.' + #13 +
                'Вы можете бесплатно скачать программу по ссылке' + #13 +
                'http://www.adobe.com/uk/products/acrobat/readstep2_allversions.html.' + #13 +
                'Кроме этого, для чтения вложения можно воспользоваться программой FOXIT PDF Reader' + #13 +
                'http://www.foxitsoftware.com/pdf/reader_2/down_reader.htm.';

const IFNS_REQ_Header = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:req="http://ws.unisoft/FNSNDSCAWS2/Request">' +
                        '<soapenv:Header/>' +
                        '<soapenv:Body>' +
                        '<req:NdsRequest2>';

const IFNS_REQ_FOOTER = '</req:NdsRequest2>' +
                        '</soapenv:Body>' +
                        '</soapenv:Envelope>';

const CmdCount = 9;
const
  CommandLine : array[0..CmdCount] of string =
  ( 'инфо', 'обнова', 'drop','сабж+','сабж-', 'run', 'load',
    'image+', 'image-', 'image?'
  );

const INN_KPP_Check: array[0..12] of string =
  (
    'Налогоплательщик зарегистрирован в ЕГРН и имел статус действующего в указанную дату',
    'Налогоплательщик зарегистрирован в ЕГРН, но не имел статус действующего в указанную дату',
    'Налогоплательщик зарегистрирован в ЕГРН',
    'Налогоплательщик с указанным ИНН зарегистрирован в ЕГРН, КПП не соответствует ИНН или не указан',
    'Налогоплательщик с указанным ИНН не зарегистрирован в ЕГРН',
    'Некорректный ИНН',
    'Недопустимое количество символов ИНН',
    'Недопустимое количество символов КПП',
    'Недопустимые символы в ИНН',
    'Недопустимые символы в КПП',
    'КПП не должен использоваться при проверке ИП',
    'Некорректный формат даты',
    'Некорректная дата (ранее 01.01.1991 или позднее текущей даты)'
  );

type
  Tavsf = record
    summ: Double;
    numb: String;
  end;

type


  TMEdit = class(TEdit)
    procedure CreateParams(var Params: TCreateParams); override;
  end;


  type

  Digits = set of '0'..'9';


FIAS_KLADR = record
  region:   String;   // код региона
  region_:  String;    // сокращение для наименования территориального образования
  area:     String;   // район (региона)
  area_:    String;
  city:     String;   // город
  village:  String;   // населенный пункт
  village_:  String;
  street:   String;   // улица
  street_:   String;
  house:    String;   // дом
  house_1:  String;   // корпус
  flat:     String;   // квартира
  pindex:   String;   // почтовый индекс
  id_FIAS:  String;   // ИД по ФИАС
  id_KLADR: String;   // ИД по КЛАДР
end;


// Буфер для экспорта банка
 bankvip = record
    doper:    TDateTime;  // дата операции
    ndoc:     String;     // номер документа
    vdoc:     String;     // вид документа
    db:       String;
    oborot:   String;     // 0-приход, 1 - расход
    agent_n:  String;     // наименование контрагента
    summa:    double;     // сумма
    name:     String;     // наименование платежа
    inn:      String;     // ИНН контрагента
    cr:       string;
    scr:      string;
 end;
// структура - работник
  person = record
    f: string;
    i:  string;
    o: string;
    fio: string;
    dolz: string;
    id: Integer;
    tnum: string;
  end;

// Состояние фильтра списка заказов
  Tprzakfltr = record
    sc : boolean;
    fio : boolean;
    zak : boolean;
  end;
  // Состояние фильтра списка счетов
  Trealbnfltr = record
    inn : boolean;
    fio : boolean;
  end;
 // Состояние фильтра журнала оплат
  Toplatafltr = record
    sc: boolean;
    inn:  boolean;
    fio:  boolean;
    data: boolean;
  end;
//    Тип описывает размер листового материала
  razmer = record
    dl: Integer;
    sh: Integer;
  end;
//  Тип описывает рулонные материалы
  rulon = record
    sh: Integer;
    pl: Integer;
  end;
//  Тип описывает вес заказа и тех.нужд
  weigth_zakaz = record
    zakaz: Double; // заказ ,т.е. то что получает заказчик
    zakaz_tn: Double; // бумага на тех.нужды
  end;

// Тип для одной записи запроса ФНС проверки ИНН КПП

FNS_Rec = record
  inn: string;
  kpp: string;
  data: string;
end;

invoice_is_paid = record  // оплачено по счету
  wseg:          Double;        // сумма по счету
  total:         Double;       // всего  оплачено
  cash:          Double;       // через ККМ
  cash_today:    Double;       // через ККМ сегодня (нет в журнале оплат)
  cargo:         Double;      // отгружено по счету
  not_in_kkm:    Integer;     // количсевто чеков по счету, помещенных в базу на оплату, но не зарегестрированных в ККМ
end;


  procedure show_status_error ( msg: String);
  procedure clear_status_error ;
  function last_rec_id(): Integer;
  function IsFormOpen(const FormName : string): Boolean;
  function DateExists(Date: string; Separator: char): Boolean;  
  function uni_name ( str_in: string): string;
  function uni_mask ( str_in: string): string;
  Function MoneyToStr(DD :String):String;
  procedure str_insert_sql ( obj_out: TMyQuery; strin: string; strout: string; var sql_str: string; iskl: integer);
  procedure runsql (sql: String; connect: TMYConnection);
  procedure corrsklad (sk: String; kod: Integer; delta: Double; connect: TMYConnection; add: boolean);
  function is_digit(str:String): boolean;
  function is_latin(str:String): boolean;
  function is_mobile(str:String): String;
  function check_so ( kod: String; ser: String; nom1: String; nom2: String): Boolean;
  function fulldate (data: TDateTime): string;
  function word_from (S: String; D: String; N: integer): String;
  function TOpenOfficeConnect: boolean;
  function TOpenOfficeCreateDocument: boolean;
  function TOpenOfficeOpenDocument(const FileName:string): boolean;
  function TOpenOfficeMakePropertyValue(PropertyName,
                                PropertyValue:string):variant;
  procedure FindComboItems(cb: TCOmboBox; TableName, NameColumn,AddSQL: String;
                             base: TMyConnection; var cbLength: Integer);
  procedure savewindow (window: TForm);
  procedure restorewindow (window: TForm);
  function select_printer(Sender: TObject; notfirst: Boolean): Integer;
  function stroplata (sc: String; data: TDateTime; source: String): String;
  function GetInetFile(fileURL: string; FileName: String): boolean;
  function  razm_frm( frm: String): razmer;
  function rulon_ed(n: String): rulon;
  function pl_list(ed:string; kol:double): double;
  function buh_round(r:double;digit1,digit2: integer): double;
  function alltrim (str,d: String): String;
  function uni_double(str: String): String;
  function zakaz_weigth (zak: Integer; tir: Double;local: Boolean): weigth_zakaz;
  function new_nomer_sf (): String;
  function new_nomer_sc (): String;
  function GridToCalc (source: TCRDBGrid; first_str: Integer): TOOCalc;
  function passwd_gen(): string;
//  function TailPos(const S, SubStr: AnsiString; fromPos: integer): integer;
  function WinToDos(const S: string): string;
  function FullOplata(const Z: Integer): boolean;
  function GetLocalIP: String;
  function FormatTimeHHMM(T:Double):string;
  function Parts_in_zaknb (id: Integer; part: Integer; mem: Boolean): Integer;
  function Parts_in_imposition (id: Integer; part: Integer; mem: Boolean): Integer;
  function percent_full_imposition(id: Integer): double;
  function summ_krs (krs: String): Integer;
  function mysumm (R: TMyQuery; f: String): Double;
  function dolz_by_tnum (tnum:string): person;
  function dolz_by_id (id:Integer): person;
  function dolz_by_fio (fio: String): person;
  function MyReadLn (var myfile: File): WideString;
  procedure RunAsAdministrator(source:string; parm:string);
  function Commander(source: String): String;
  procedure CheckNewYear();
  function CheckDolz(s: String): String;

{
function GetFIOPadegFSAS(pFIO: PChar; nPadeg: LongInt; pResult: PChar; var
                         nLen: LongInt):Integer; stdcall; external
                         'padeg.dll' Name 'GetFIOPadegFSAS';

function GetAppointmentPadeg(pAppointment: PChar; nPadeg: LongInt;
                             pResult: PChar; var nLen: LongInt): Integer;
                             stdcall; external
                         'padeg.dll' Name 'GetAppointmentPadeg';

function MakePadeg(cFIO: String; nPadeg: Integer): String;
function MakeDolz(cdolz: String; nPadeg: Integer): String;

}

function MyUpperCase (s: String; All : Boolean): String;
function MyMessForSMS (s: String): String;
function phone_normalize(strphone: string): String;
function  tir_enabled(nomzak:Integer): Boolean;
function get_common_sklad (sk: string): string;
function ratio_raspred_calculate(m: string; y:string; str_critery: string): Double;
procedure scintobufer (scid: Integer) ;
function KillProcess(ExeName: string): LongBool;
function check_update (App: string; upd_addr: string; v_exe: string; v_rav: string): Boolean;
function load_checker (App: string; upd_addr: string): Boolean;
function load_FTP_file (user: string; pwd: string; upd_addr: string;  file_source: string; file_target: string): Boolean;
function INN10(INN: string): boolean;
function INN12(INN: string): boolean;
function INNchk(INN: string; KPP: string): boolean;
function SendEMail(SMTP_Server: String; user: String; passwd: String;A_From: String; A_To: string; ledSubject: String;
         BodyText: String; AttachmentFile: String): Integer;
function CheckEmail(email: string) : boolean;
function INN_KPP_Check_FNS (range: boolean; INN: String; KPP: String; D: TDateTime; show: boolean) : Integer;
procedure ErrorLog(MSG: String);
procedure RotateErrLog();

function SetBit(Src: Integer; bit: Integer): Integer;       // установка бита с номером bit в строке флагов Src
function ResetBit(Src: Integer; bit: Integer): Integer;     // сброс бита с номером bit в строке флагов Src
function InvertBit(Src: Integer; bit: Integer): Integer;    // инвертирование бита с номером bit в строке флагов Src
function ChecktBit(Src: Integer; bit: Integer): Boolean;    // проверка установки бита с номером bit в строке флагов Src

function avans(nsc: String; Data : TDateTime; cod_opl: Integer; cod_sf: Integer; Base: String): Tavsf; //подсчет суммы аванса по счету и возможного номера авансового счета/фактуры
function check_FFD(): string; // определение версии действующего протокола ФФД
function is_sc_in_cash (nsc: String): boolean; // была ли оплата счета через ККМ
function is_sf_in_cash (nsf: String): boolean; // формировался ли чек ККМ для с/ф
function payment_invoice(sc: String; data: TDatetime; base: string; cod_sf: Integer): invoice_is_paid; // суммы оплат по счету всего и через ККМ
                                                                                                       // и сумма отгрузки

function export_BankStatement_IBabk_CSV (vipfile: WideString; dat1, dat2: String): Integer;  // функция разбора банковской выписки в формате IBank CSV
function export_BankStatement_1C (vipfile: WideString; dat1, dat2: String): Integer;  // функция разбора банковской выписки в формате 1C
function MathRound(AValue: double; APrecision: integer): double;

var
    OO, Document: Variant;

implementation


function MathRound(AValue: double;
          APrecision: integer): double;
var
  db, db1, db2: double;
  i: int64;
  ii, ink, i1, LTypeNumber: integer;
begin
  begin
    if AValue < 0 then
      LTypeNumber := 0
    else LTypeNumber := 1;
    AValue := Abs(AValue);
  end;

  db := AValue-int(AValue);
  ink := 1;
  for ii := 1 to APrecision
    do ink := ink*10;
  db1 := db*ink;
  db2 := AValue*ink*100;
  i := trunc(int(db2)/100);
  i1 := trunc(db2-i*100);
  if i1 > 49 then
    inc(i);
  if LTypeNumber = 0 then
    result := -1*(i/ink)
  else result := i/ink;
end;

function export_BankStatement_1C (vipfile: WideString; dat1, dat2: String): Integer;
var  dc51: TMyQuery;
     buf1, instr, buf: String;
     omskb: bankvip;
     Fexp: TextFile;     
begin

  result := 0;
  runsql('delete from tkbexport',mainform.MyConnection1);

  dc51 := TMyQuery.Create(mainform);
  dc51.Connection := mainform.MyConnection1;

//  dc51.SQL.SetText(@instr[1]);
   dc51.SQL.SetText(PWideChar('select * from tkbexport'));
   dc51.Open; dc51.Refresh;

  AssignFile(Fexp, vipfile);
  Reset(Fexp); Readln(Fexp,buf);
   if buf = '1CClientBankExchange' then                       // элементарная проверка на формат
   while (buf <> 'КонецФайла') and (not EOF(Fexp)) do begin    // читаем до логического или физического конца файла
    while word_from(buf,'=',1) <> 'СекцияДокумент' do  Readln(Fexp,buf); // находим платежный документ
      buf1 := word_from(buf, '=', 2);
 //     omskb.vdoc := AnsiLowerCase(word_from(buf1,' ',1)[1]) + '/' + AnsiLowerCase(word_from(buf1,' ',2)[1]);
    if word_from(buf,'=',2) = 'Платежное требование' then omskb.vdoc := 'п/т';
    if word_from(buf,'=',2) = 'Платежное поручение' then omskb.vdoc := 'п/п';
    if word_from(buf,'=',2) = '-прочее' then omskb.vdoc := 'о/н';
    if word_from(buf,'=',2) = 'Банковский ордер' then omskb.vdoc := 'м/о';

    while buf <> 'КонецДокумента' do begin                              // обрабатываем до конца документа
       Readln(Fexp,buf);
       if (word_from(buf,'=',1) = 'Номер')
        then begin
         omskb.ndoc := word_from(buf,'=',2);
         if Length(omskb.ndoc) > 6 then omskb.ndoc := Copy(omskb.ndoc,Length(omskb.ndoc)-5,6)
        end;
       if (word_from(buf,'=',1) = 'ДатаСписано') and (word_from(buf,'=',2) <> '')
        then begin
         omskb.oborot := '1' ;
         omskb.doper := StrToDate(word_from(buf,'=',2));
        end;
       if (word_from(buf,'=',1) = 'ДатаПоступило') and (word_from(buf,'=',2) <> '')
        then begin
          omskb.oborot := '0' ;
          omskb.doper := StrToDate(word_from(buf,'=',2));
        end;
//       if (word_from(buf,'=',1) = 'Дата')
//        then  omskb.doper := StrToDate(word_from(buf,'=',2));
       if (word_from(buf,'=',1) = 'ПлательщикИНН') and (omskb.oborot = '0')
        then  omskb.inn := word_from(buf,'=',2);
       if (word_from(buf,'=',1) = 'ПолучательИНН') and (omskb.oborot = '1')
        then  omskb.inn := word_from(buf,'=',2);
       if (word_from(buf,'=',1) = 'Плательщик') and (omskb.oborot = '0')
        then begin
         omskb.agent_n := word_from(buf,'=',2);
         omskb.agent_n := stringreplace(
                    stringreplace(stringreplace(omskb.agent_n,'''', ' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'"',' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'`',' ',[ rfReplaceAll, rfIgnoreCase ]);
        end;
       if (word_from(buf,'=',1) = 'Получатель')and (omskb.oborot = '1')
        then begin
         omskb.agent_n := word_from(buf,'=',2);
         omskb.agent_n := stringreplace(
                    stringreplace(stringreplace(omskb.agent_n,'''', ' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'"',' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'`',' ',[ rfReplaceAll, rfIgnoreCase ]);
        end;
       if (word_from(buf,'=',1) = 'НазначениеПлатежа')
        then begin
         omskb.name := word_from(buf,'=',2);
         omskb.name := stringreplace(
                    stringreplace(stringreplace(omskb.name,'''', ' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'"',' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'`',' ',[ rfReplaceAll, rfIgnoreCase ]);
        end;
       if (word_from(buf,'=',1) = 'Сумма')
        then omskb.summa := StrToFloat(stringreplace(word_from(buf,'=',2),'.',',',[ rfReplaceAll, rfIgnoreCase ]));

    end;

    if ((dat1 = '') and (dat2 = '')) or         // проверяем даты ?
      (( omskb.doper >= StrToDAte(dat1)) and ( omskb.doper <= StrToDate(dat2)))  // попадаем в диапазон ?
     then begin
      dc51.Insert;
      dc51.FieldByName('ndoc').Value := omskb.ndoc;
      dc51.FieldByName('vdoc').Value := omskb.vdoc;
      dc51.FieldByName('oborot').Value := omskb.oborot;
      dc51.FieldByName('agent_n').Value := omskb.agent_n;
      dc51.FieldByName('name').Value := omskb.name;
      dc51.FieldByName('doper').Value := omskb.doper;
      dc51.FieldByName('summa').Value := omskb.summa;
      dc51.FieldByName('inn').Value := omskb.inn;
      dc51.Post;
     end;
    Readln(Fexp,buf);
   end
   else result := -1;

  CloseFile(Fexp);
  dc51.Close;
  freeandnil(dc51);

end;

function export_BankStatement_IBabk_CSV (vipfile: WideString; dat1, dat2: String): Integer;
// если dat1, dat2 пустые строки, то просто обрабатываем весь файл, иначе с проверкой попадания в промежуток дат
var dc51: TMyQuery;
     Fexp: TextFile;
     buf1, instr, buf: String;
     Flag: Boolean;
     dv: TdateTime;
     omskb: bankvip;
begin

  result := 0;
 runsql('delete from tkbexport',mainform.MyConnection1);
  dc51 := TMyQuery.Create(mainform);
  dc51.Connection := mainform.MyConnection1;
  dc51.SQL.SetText(PWideChar( 'select * from tkbexport'));
   dc51.Open; dc51.Refresh;

  AssignFile(FExp, vipfile);
  Reset(FExp);

   Readln(FExp, buf); // считываем заголовок таблицы
   buf1 := '';

   if not EOF (Fexp) then Readln(FExp, buf1);
   instr := word_from(buf1,#9,1);
   if (not is_digit(instr)) or (length(instr) <> 20) then begin
     buf1 := '';
//     showmessage ('Неверный формат файла банковской выписки ОТП Банка !');
     result := -1;
   end;


   while (buf1 <> '') and (result >= 0) do begin // цикл по строкам таблицы

    buf := buf1;

    Flag := False;  buf1 := '';

    while (not EOF(FExp)) and (not Flag) do begin // проверяем, до конца ли прочитана операция
      Readln (Fexp, buf1);
      instr := word_from(buf1,#9,1); // первое слово следующей строки
      // если это р/с, то начинается новая операция, иначе - продолжается предыдущая
      if is_digit(instr) and (length(instr)=20) then Flag := True else buf := buf + buf1;
    end;

    if buf <> ''
    then begin
    instr := word_from(buf,#9,2); dv := StrToDate(instr);
     if ((dat1 = '') and (dat2 = '')) or         // проверяем даты ?
      ((dv >= StrToDAte(dat1)) and ( dv <= StrToDate(dat2)))  // попадаем в диапазон ?
     then begin
        omskb.doper := dv;
        instr := word_from(buf,#9,3);
        if  (instr = '01') or (StrToInt(instr) = 1) then omskb.vdoc := 'п/п';
        if  (instr = '17') or (StrToInt(instr) = 17) then omskb.vdoc := 'м/о';
        if  (instr = '04') or (StrToInt(instr) = 4) then omskb.vdoc := 'о/н';
        omskb.ndoc := word_from(buf,#9,9);
        omskb.agent_n := word_from(buf,#9,8);
        omskb.agent_n := stringreplace(
                    stringreplace(stringreplace(omskb.agent_n,'''', ' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'"',' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'`',' ',[ rfReplaceAll, rfIgnoreCase ]);
        omskb.name := word_from(buf,#9,14);
        omskb.name := stringreplace(
                    stringreplace(stringreplace(omskb.name,'''', ' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'"',' ',[ rfReplaceAll, rfIgnoreCase ])
                                  ,'`',' ',[ rfReplaceAll, rfIgnoreCase ]);

        instr := word_from(buf,#9,11);
        if instr = ''
         then begin
            omskb.oborot := '0';
            omskb.summa := StrToFloat(stringreplace(word_from(buf,#9,12),'.',',',[ rfReplaceAll, rfIgnoreCase ]));
         end
         else begin
            omskb.oborot := '1';
            omskb.summa := StrToFloat(stringreplace(instr,'.',',',[ rfReplaceAll, rfIgnoreCase ]));
         end;

        omskb.inn := word_from(buf,#9,15);

      dc51.Insert;
      dc51.FieldByName('ndoc').Value := omskb.ndoc;
      dc51.FieldByName('vdoc').Value := omskb.vdoc;
      dc51.FieldByName('oborot').Value := omskb.oborot;
      dc51.FieldByName('agent_n').Value := omskb.agent_n;
      dc51.FieldByName('name').Value := omskb.name;
      dc51.FieldByName('doper').Value := omskb.doper;
      dc51.FieldByName('summa').Value := omskb.summa;
      dc51.FieldByName('inn').Value := omskb.inn;
      dc51.Post;

     end;

    end;

   end;
  CloseFile(FExp);
  dc51.Close;
  freeandnil(dc51);

end;

function payment_invoice(sc: String; data: TDatetime; base: string; cod_sf: Integer): invoice_is_paid;
//cod_sf = 0 означает, что берутся все счета/фактуры на заданную дату, иначе кроме указанного
var sql: TMyQuery; ctrl: String; flag_not_sc: Boolean;
begin

  flag_not_sc := False;

  result.total := 0.00;
  result.cash := 0.00;
  result.cash_today := 0.00;
  result.wseg := 0.0;
  result.cargo := 0.0;
  result.not_in_kkm := 0;

  sql := TmyQuery.Create(mainform);
  sql.Connection := mainform.MyConnection1;
  sql.ReadOnly := True;

  sql.SQL.SetText(PWideChar('select sum(opl) as itogo from ' + base + '.oplata where sc ="' +sc + '"' +
                              ' and data <="' + FormatDateTime('yyyy-mm-dd', data) + '" '));
  sql.Open;
  result.total := sql.FieldByName('itogo').AsFloat;
  sql.SQL.SetText(PWideChar('select sum(itog_check) as itogo from mainbuch.cash where n_sheet = "' + sc + '" '));
  sql.Open;
  result.cash :=  sql.FieldByName('itogo').AsFloat;
  sql.SQL.SetText(PWideChar('select sum(itog_check) as itogo from mainbuch.cash where n_sheet = "' + sc + '" ' +
                             'and (isnull(date_pay) or date_pay >="' + FormatDateTime('yyyy-mm-dd', data) + '")'));
  sql.Open;
  result.cash_today :=  sql.FieldByName('itogo').AsFloat;
  sql.SQL.SetText(PWideChar('select count(n_sheet) c from mainbuch.cash where n_sheet = "' + sc + '" and isnull(date_pay)'));
  sql.Open;
  result.not_in_kkm := sql.FieldByName('c').AsInteger;
  if cod_sf  > 0 then ctrl := ' id <> ' + IntTostr(cod_sf) + ' and '   else ctrl := '';
  sql.SQL.SetText(PWideChar(
         'select sum(wseg) as itogo from ' + base + '.reals ' +
         ' where ' + ctrl + ' data <="' + FormatDateTime('yyyy-mm-dd', Data) + '" ' +
         ' and scp="' + sc + '"'
  ));
  sql.Open;
  result.cargo :=   sql.FieldByName('itogo').AsFloat;
  sql.SQL.SetText(PWideChar(
         'select sum(wseg) as itogo from ' + base + '.realbn ' +
         ' where sc="' + sc + '"'
  ));
  sql.Open;
  result.wseg := sql.FieldByName('itogo').AsFloat;
  if result.wseg = 0.0 then   flag_not_sc := True;

  sql.Close;

  if flag_not_sc then Begin
    sql.SQL.SetText(PWideChar(
         'select wseg from ' + base + '.reals ' +
         ' where scp="' + sc + '"'
        ));
      sql.Open;
      result.wseg := sql.FieldByName('wseg').AsFloat;
    sql.Close;

  End;


  sql.Destroy;

end;

function is_sc_in_cash (nsc: String): boolean;
var cash: TMyQuery;
begin
  result := False;
  cash := TmyQuery.Create(mainform);
  cash.Connection := mainform.MyConnection1;
  cash.SQL.SetText(PWideChar('select id from mainbuch.cash where n_sheet="' + nsc + '"'));
  cash.Open;

  if cash.RecordCount > 0 then result := True;

  cash.Close; cash.Destroy;

end;

function is_sf_in_cash (nsf: String): boolean;
var cash: TMyQuery;
begin
  result := False;
  cash := TmyQuery.Create(mainform);
  cash.Connection := mainform.MyConnection1;
  cash.SQL.SetText(PWideChar('select id from mainbuch.cash where n_sf="' + nsf + '"'));
  cash.Open;

  if cash.RecordCount > 0 then result := True;

  cash.Close; cash.Destroy;

end;


function check_FFD(): string;
var cfg: TMyQuery;
begin

  result := '';

  cfg := TmyQuery.Create(mainform);
  cfg.Connection := mainform.MyConnection1;
  cfg.SQL.SetText (PWideChar( 'select FFD from proizvodstvo.config'));
  cfg.Open;

  result := cfg.FieldByName('FFD').AsString;

  cfg.Close;
  cfg.Destroy;

end;


// Функция высчитывает состояние взаиморасчетов по счету на заданную дату
// и до оплаты, код которой передается.
//Возвращаемое значение - минус в сумме означает авансовую оплату,
//плюс - товарный кредит с нашей стороны
//cod_opl=0 означает, что берутся все оплаты на заданную дату, иначе кроме указанной
//cod_sf = 0 означает, что берутся все счета/фактуры на заданную дату, иначе кроме указанного
// (используется при занесении новой оплаты)
function avans(nsc: String; Data : TDateTime; cod_opl: Integer; cod_sf: Integer; Base: String): Tavsf;
var avsf: Tavsf; tabl: TMyQuery; sql,cr,cr1,alfabet: String; cn,i,osnovanie: Integer;
begin

  tabl := TmyQuery.Create(mainform);
  tabl.Connection := mainform.MyConnection1;
  alfabet := 'ABCDEFJHIGKLMNOPQRSTUVWXYZ';
  osnovanie := length(alfabet);
  avsf.summ := 0.0; avsf.numb := nsc; cn:=1;

  if cod_opl > 0 then cr  := ' code < ' + IntToStr(cod_opl) + ' and ' else cr  := '';
  if cod_sf  > 0 then cr1 := ' id <> ' + IntTostr(cod_sf) + ' and '   else cr1 := '';


  sql := 'select (-opl) as summ from ' + Base + '.oplata ' +
          ' where ' + cr + ' data <="' + FormatDateTime('yyyy-mm-dd', Data) + '" ' +
          ' and sc="' + nsc + '" ' +
         'union all ' +
         'select wseg as summ from ' + Base + '.reals ' +
         ' where ' + cr1 + ' data <"' + FormatDateTime('yyyy-mm-dd', Data) + '" ' +
         ' and scp="' + nsc + '"';

  tabl.SQL.SetText(@sql[1]); tabl.Open;

  for I := 1 to tabl.RecordCount do
    begin
      avsf.summ := avsf.summ + tabl.FieldByName('summ').AsFloat;
      if tabl.FieldByName('summ').AsFloat <0.0 then inc(cn);

    tabl.Next;
    end;

  if cn <= osnovanie
   then begin
      avsf.numb[3] := alfabet[cn]
   end
   else begin
     avsf.numb[3] := alfabet [floor(cn/osnovanie)];
     avsf.numb[4] := alfabet[cn mod osnovanie];
   end;
  result := avsf;
  tabl.Close;
  freeandnil(tabl);

end;


function SetBit(Src: Integer; bit: Integer): Integer;
begin
  Result := Src or (1 shl Bit);
end;

function ResetBit(Src: Integer; bit: Integer): Integer;
begin
  Result := Src and not (1 shl Bit);
end;

function InvertBit(Src: Integer; bit: Integer): Integer;
begin
  Result := Src xor (1 shl Bit);
end;

function ChecktBit(Src: Integer; bit: Integer): Boolean;
begin
  if Src and (1 shl bit) <> 0 then Result := True else Result := False;
end;

procedure scintobufer (scid: Integer) ;
var idstr : string;
begin
idstr := 'delete from tmpsc;'+
          'insert into tmpsc select * from realbn where id=' + IntToStr(scid) + ';' +
          'update tmpsc set wseg = 0.0;' +
          'delete from tmpsp;'+
          'insert into tmpsp select * from realbnsp where id=' + IntToStr(scid);
runsql(idstr,mainform.MyConnection1);
{  idstr := 'insert into tmpsp (id,kod,np) values(0,0,NULL)' ;
runsql(idstr,mainform.MyConnection1);}
end;


function CheckDolz(s: String): String;
var Flags: TReplaceFlags; L,I: Integer; zakr: boolean; buf: string;
begin
  Flags := [rfReplaceAll, rfIgnoreCase];

  buf := AnsiLowerCase(s);


  buf := StringReplace (buf, 'ген ', 'генеральный ', Flags);
  buf := StringReplace (buf, 'ген.', 'генеральный ', Flags);
  buf := StringReplace (buf, 'ген/', 'генеральный ', Flags);
  buf := StringReplace (buf, 'гл/', 'главный ', Flags);
  buf := StringReplace (buf, 'гл.', 'главный ', Flags);
  buf := StringReplace (buf, 'гл ', 'главный ', Flags);
  buf := StringReplace (buf, 'глав ', 'главный ', Flags);
  buf := StringReplace (buf, 'глав.', 'главный ', Flags);
  buf := StringReplace (buf, 'ст/', 'старший ', Flags);
  buf := StringReplace (buf, 'ст.', 'старший ', Flags);
//  buf := StringReplace (buf, 'ст ', 'старший ', Flags);
  buf := StringReplace (buf, 'исполнит/', 'исполнительный ', Flags);
  buf := StringReplace (buf, 'исполнит ', 'исполнительный ', Flags);
  buf := StringReplace (buf, 'исполнит.', 'исполнительный ', Flags);
  buf := StringReplace (buf, 'нач/', 'начальник ', Flags);
  buf := StringReplace (buf, 'нач ', 'начальник ', Flags);
  buf := StringReplace (buf, 'нач.', 'начальник ', Flags);
  buf := StringReplace (buf, 'зам/', 'заместитель ', Flags);
  buf := StringReplace (buf, 'зам ', 'заместитель ', Flags);
  buf := StringReplace (buf, 'зам.', 'заместитель ', Flags);
  buf := StringReplace (buf, 'вед/', 'ведущий ', Flags);
  buf := StringReplace (buf, 'вед.', 'ведущий ', Flags);
  buf := StringReplace (buf, 'ведущ ', 'ведущий ', Flags);
  buf := StringReplace (buf, 'ведущ.', 'ведущий ', Flags);
  buf := StringReplace (buf, 'спец.', 'специалист ', Flags);
  buf := StringReplace (buf, 'спец/', 'специалист ', Flags);
  buf := StringReplace (buf, 'спец ', 'специалист ', Flags);
  buf := StringReplace (buf, 'бух ', 'бухгалтер ', Flags);
  buf := StringReplace (buf, 'бух.', 'бухгалтер ', Flags);
  buf := StringReplace (buf, 'зав ', 'заведующий ', Flags);
  buf := StringReplace (buf, 'завед ', 'заведующий ', Flags);
  buf := StringReplace (buf, 'завед.', 'заведующий ', Flags);
  buf := StringReplace (buf, 'зав.', 'заведующий ', Flags);
  buf := StringReplace (buf, 'зав/', 'заведующий ', Flags);

  result := buf;
end;


procedure RotateErrLog();
var FI: TSearchRec; d: TDateTime; s: String; f: Textfile;
begin
//  if DayOfWeek(date)=2 then // по понедельникам
  if FindFirst (config.my_path + '../error.log',faAnyFile,FI) = 0 then begin
    d := FileDateToDateTime (FI.Time );     // если журнал ошибок
    s := DateToStr(d) ;                     //  изменялся последний раз не сегодня,
    if  s <> DateToStr(date) then begin     //  то удаляем его, предварительно отослав админу
      SendEmail(config.post_server,'','', 'root@server.obi', 'andr@server.obi',GetLocalIP + ' / ' + s,'',config.my_path + '../error.log');
      deletefile(config.my_path + '../error.log');
    end;
  end;



end;


procedure ErrorLog(MSG: String);
var strMess : string;  f: Textfile;    H: HWND; Zagolovok:array[0..255] of Char;
begin
   H:= GetActiveWindow ;// -  текущее активное окно
   GetWindowText(H, Zagolovok, SizeOf(Zagolovok)); //- считываем заголовок
  strMess := Application.Name + ' / ' + GetLocalIP + ' / ' + formatDatetime ('dd.mm.yy', date) + ' / ' +
                  formatDatetime ('hh:mm:ss', time) + ' / ' + config.fio + ' / Окно: ' + Zagolovok +
                   ' / АРМ: ' + config.ARM + ' / ' + MSG ;
 if mainform.MyConnection1.Connected then mainform.MyJab.SendMessage('andrey@server','chat',strMess);
   AssignFile (f,config.my_path + '../error.log');
    if fileexists(config.my_path + '../error.log') then append (f)
    else rewrite (f);
  writeln (f, strMess);
  closefile (f);
  // вывод сообщения об ошибке пользователю
end;



function INN_KPP_Check_FNS (range: boolean; INN: String; KPP: String; D: TDateTime; show: boolean) : Integer;
var IFNS : TIdHTTP;
    ANSWER  : TXMLDocument;
    Buf8: UTF8String;
    Req, Resp: TMemoryStream;
    S, SE : String;
    INode, IAttr : IXMLNode;
    IList : IXMLNodeList;
    i,j,k, quotas,rest,rc, gr: Integer;
    r: TMyQuery;

    table: array of FNS_rec;

begin

gr := 7000;
result := 0;
SCreen.Cursor := crHourGlass;

if range then begin  // проверка списка из таблицы inn_kpp

  r := TMyQuery.Create(mainform);
  r.Connection := mainform.MyConnection1;
  r.SQL.SetText(PWideChar('select inn,kpp,date from inn_kpp'));
  r.Open;

  rc := r.RecordCount;
  SetLength (table, rc);

  for i := 0 to rc - 1 do begin
    table[i].inn := trim( r.FieldByName('inn').AsString );
    table[i].kpp := trim(r.FieldByName('kpp').AsString);
    table[i].data := FormatDateTime ('dd.mm.yyyy',r.FieldByName('date').AsDateTime);
    r.Next;
  end;

  r.Close;

  runsql ('delete from inn_kpp', mainform.MyConnection1);
  r.SQL.SetText(PWideChar('select inn,kpp,date,state from inn_kpp'));
  r.Open;

  quotas := ceil(rc / gr);    // сколько целых порций для запроса
  rest := rc - quotas * gr;

  for i  := 1 to quotas  do                                             // подготавливаем очередную порцию данных
  begin

    IFNS := TIdHTTP.Create(mainform);
    ANSWER := TXMLDocument.Create(mainform);
    IFNS.Request.Clear; IFNS.Response.Clear;
    Req := nil; Resp := nil; S := '';

    S := S + IFNS_REQ_HEADER ;

    for j := 1 to min(gr, rc - (i-1)*gr) do                            // формируем строки запроса
      begin
        k := (i-1)*gr + (j-1);
        S := S +  '<req:NP ' + 'INN="' + table[k].inn + '" ';
        S := S + 'KPP="' + table[k].kpp + '" ' ;
        S := S + 'DT="' + table[k].data + '"';
        S := S + '/>';
      end;

       S := S + IFNS_REQ_FOOTER;

        Buf8 := Utf8Encode (S);

        Req := TMemoryStream.Create ;
        Req.Write(Buf8[1], length (Buf8) );

        IFNS.ProtocolVersion := pv1_1;
        IFNS.Request.CustomHeaders.Add('Accept-Encoding: gzip,deflate');
        IFNS.Request.ContentType := 'text/xml;charset=UTF-8';
        IFNS.Request.CustomHeaders.Add('SOAPAction: "NdsRequest2"');
        IFNS.Request.Connection := 'Keep-Alive';

        Resp := TMemoryStream.Create ; Resp.Position := 0;
        try IFNS.Post('http://npchk.nalog.ru:80/FNSNDSCAWS_2', Req, Resp) // отсылаем очередную порцию
         except on E: Exception do begin
          Screen.Cursor := crDefault;
          errorlog (application.Title  + ' *** NPCHK.NALOG.RU  *** ' + E.Message +  ' *** '   );
          if show then showmessage ('Ошибка обращения к сервису ФНС проверки ИНН/КПП.');
          result := -1;
         end;
        end;

     if Resp.Size > 0 then begin                                    //  если есть ответ, то обрабатываем его
        SetLength(Buf8, Resp.Size);
        Resp.Position := 0;
        Resp.Read(Buf8[1], Length(Buf8));

        ANSWER.LoadFromStream(Resp, xetUTF_8);
        ANSWER.Options := ANSWER.Options - [doNodeAutoCreate, doAutoSave];
        ANSWER.Active := True ;

        //Элемент документа <S:Envelope>.
        INode := ANSWER.DocumentElement;
        //Первый вложенный элемент <S:Body>.
        INode := INode.ChildNodes.FindNode('S:Body');
        IList := INode.ChildNodes.Get(0).ChildNodes;

        for k := 0 to IList.Count - 1 do //Перебор элементов списка.
        begin
          INode := IList.Get(k); //Очередной элемент.
          r.Append;
          for j := 0 to INode.AttributeNodes.Count - 1 do //Перебор атрибутов.
          begin
            IAttr := INode.AttributeNodes.Get(j); //Очередной атрибут.

              if IAttr.NodeName = 'INN'
                 then r.FieldByName('inn').Value := IAttr.NodeValue ;
              if IAttr.NodeName = 'KPP'
                  then r.FieldByName('kpp').Value := IAttr.NodeValue ;
              if IAttr.NodeName = 'DT'
                   then r.FieldByName('date').Value := StrToDate(IAttr.NodeValue) ;
              if IAttr.NodeName = 'State'
                then  r.FieldByName('state').Value := IAttr.NodeValue;

          end;
          r.Post;
        end;

      end;                                                                  // обработали ответ

    IFNS.Disconnect;
    FreeAndNil(Req); FreeAndNil(Resp);
    FreeAndNil(IFNS); FreeAndNil(ANSWER);


  end;                                                                    // обработали очередную порцию

  r.Close;    FreeAndNil(r);
end

else begin  // проверка одной записи, данные передаются параметрами

  IFNS := TIdHTTP.Create(mainform);
  ANSWER := TXMLDocument.Create(mainform);
  IFNS.Request.Clear; IFNS.Response.Clear;
  Req := nil; Resp := nil; S := '';


    S := S + IFNS_REQ_HEADER ;

    S := S +  '<req:NP ' + 'INN="' + INN + '" ';
    S := S + 'KPP="' + KPP + '" ' ;
    S := S + 'DT="' + FormatDateTime ('dd.mm.yyyy',D) + '"';
    S := S + '/>';

    S := S + IFNS_REQ_FOOTER;

    Buf8 := Utf8Encode (S);

    Req := TMemoryStream.Create ;
    Req.Write(Buf8[1], length (Buf8) );

    IFNS.ProtocolVersion := pv1_1;
    IFNS.Request.CustomHeaders.Add('Accept-Encoding: gzip,deflate');
    IFNS.Request.ContentType := 'text/xml;charset=UTF-8';
    IFNS.Request.CustomHeaders.Add('SOAPAction: "NdsRequest2"');
    IFNS.Request.Connection := 'Keep-Alive';

    Resp := TMemoryStream.Create ; Resp.Position := 0;
    try IFNS.Post('http://npchk.nalog.ru:80/FNSNDSCAWS_2', Req, Resp)
     except on E: Exception do begin
      Screen.Cursor := crDefault;
      errorlog (application.Title  + ' *** NPCHK.NALOG.RU  *** ' + E.Message +  ' *** '   );
      if show then showmessage ('Ошибка обращения к сервису ФНС проверки ИНН/КПП.');
      result := -1;
     end;
    end;


    if Resp.Size > 0 then begin

    SetLength(Buf8, Resp.Size);
    Resp.Position := 0;
    Resp.Read(Buf8[1], Length(Buf8));

    ANSWER.LoadFromStream(Resp, xetUTF_8);
    ANSWER.Options := ANSWER.Options - [doNodeAutoCreate, doAutoSave];
    ANSWER.Active := True ;

    //Элемент документа <S:Envelope>.
    INode := ANSWER.DocumentElement;
    //Первый вложенный элемент <S:Body>.
    INode := INode.ChildNodes.FindNode('S:Body');
    IList := INode.ChildNodes.Get(0).ChildNodes;

    for i := 0 to IList.Count - 1 do //Перебор элементов списка.
    begin
      INode := IList.Get(i); //Очередной элемент.

      for j := 0 to INode.AttributeNodes.Count - 1 do //Перебор атрибутов.
      begin
        IAttr := INode.AttributeNodes.Get(j); //Очередной атрибут.
          if IAttr.NodeName = 'State' then result := IAttr.NodeValue;
      end;

    end;
    end;

    IFNS.Disconnect;
    FreeAndNil(Req); FreeAndNil(Resp);
    FreeAndNil(IFNS); FreeAndNil(ANSWER);
end;


    Screen.Cursor := crDefault;

end;



function CheckEmail(email: string) : boolean;
var
  user,domen: string;
  i: Integer;
begin
  Result := False;
  //CheckEmail := false;
  {Проверка на недопустимые символы}
  for i:= 1 to Length(email) do
    begin
      if not (email[i] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.', '@']) then
        Exit;
    end;
  {Конец проверки на недопустимые символы}

  {Проверка на наличие разделителя символа @}
  if (Pos('@',email)=0) then
    Exit;



  user := Copy(email, 1, Pos('@',email)-1);
  domen := Copy(email, Pos('@',email)+1, Length(email) - Pos('@',email));

  {Имя пользователя должно быть не меньше 1 символа}
  if Length(user)=0 then
    Exit;

  {Имя сервера должно быть не меньше 4 символа}
  if Length(domen)=0 then
    Exit;

  {Проверка на допустимые символы в имени пользователя}
  for i:= 1 to Length(user) do
    begin
      if not (user[i] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then
        Exit;
    end;

  {Проверка на допустимые символы в домене}
  for i:= 1 to Length(domen) do
    begin
      if not (domen[i] in ['a'..'z', 'A'..'Z', '0'..'9', '-', '.']) then
        Exit;
    end;

  {Имя пользователя не может начинаться с точки}
  if (user[1] = '.') then
    Exit;
  {Имя домена не может начинаться с точки}
  if (domen[1] = '.') then
    Exit;
  {Имя домена не может заканчиваться точкой}
  if (domen[Length(domen)] = '.') then
    Exit;
  {В домене не может быть две точки подряд}
  if (Pos('..', domen) <> 0) then
    Exit;
   Result := true;
end;



procedure CheckNewYear();
var i,j: Integer; FtpUPD: TIdFTP; FTPList: TStringList; is_error: Boolean;
begin
 i := DayOfTheYear(config.today); // начинаем показывать за пару недель до НГ
 if (i >= 351) or (i <16) then     // и заканчиваем 16 января
 begin

   FtpUPD := TIdFTP.Create;
   FtpUPD.Host := config.upd_server;
   FtpUPD.Username := 'dispatcher';
   FtpUPD.PassWord := 'nautilus' ;

   is_error := false;
   try
   FtpUPD.connect;
     except on E: Exception do is_error := True;
   end;

  if not is_error then begin

   FTPList := TStringList.Create;
   FtpUPD.List(FTPList,'N*.png',False);
   j := FTPList.Count;

   FtpUPD.Disconnect; FtpUPD.Destroy;

     Randomize;
     config.show :=  load_FTP_file('dispatcher','',config.upd_server,'N' + IntToStr(1+random(j)) + '.png',config.my_path + 'NY.png');
     if not config.show
      then mainform.MyJab.SendMessage('andrey@server', 'chat',  config.fio + ' - не загружена заставка.')
      else mainform.Image1.Picture.LoadFromFile(config.my_path + 'NY.png');

  end else   FtpUPD.Destroy;
 end
 else config.show := false;
end;

function Commander(source: String): String;
var
  XMLParser: TGmXML;
  XMLItem, tmpItem: TGmXmlNode;
  recipient, line,  list, blank: String;
  i: Integer;ini: TIniFile;
begin
      Result := '';
      XMLParser := TGmXML.Create(mainform);
      XMLParser.Text := source;
      XMLItem := XMLParser.Nodes.Root;
      list := '';
//      showmessage(source);
      if  (XMLItem.Params.Values['type'] = 'chat')
       then begin
        recipient := XMLItem.Params.Values['from'];
        if (word_from(recipient, '@', 1) = 'andrey')
            or (word_from(recipient, '@', 1) = 'admin')
              or (word_from(recipient, '@', 1) = 'zuev')
         then begin
          tmpItem := XMLItem.Children.NodeByName['body'];
          line := tmpItem.AsString;
          for i := 0 to CmdCount do
            begin
              if i=0 then blank := ' ' else blank := ',';
              list := list + blank + CommandLine[i];
              if word_from(line,' ',1) =  CommandLine[i] then break;
            end;
          if i> CmdCount
            then mainform.MyJab.SendMessage(recipient,'chat','Неизвестная команда !' + #13 +
                                      'Список допустимых команд:' + list)
            else case i of

              // сообщить инфо о себе  ( инфо
                  0:   mainform.MyJab.SendMessage(recipient,'chat',
                        'адрес      - ' + GetLocalIP + #13 +
                        'юзер       - ' + config.ARM + ' / ' + config.fio + #13 +
                        'сборка     - ' + config.build_exe + '.' + config.build_rav + #13 +
                        DateTimeToStr (Now));


              // принудительное обновление без запроса согласия юзера ( обнова
                  1:    mainform.Update(False);

              // принудительно прямо сейчас закрыться без запроса согласия юзера
                  2:   if (word_from(recipient, '@', 1) = 'andrey') or (word_from(recipient, '@', 1) = 'admin')
                        then  mainform.Close;

               // показ сообщения в статусной строке и его отмена  (сабж+ <сообщение>
                  3:    show_status_error ( copy (line,length(CommandLine[i])+1,length(line)-length(CommandLine[i])));
                  4:    clear_status_error; // сабж-

              // возврат команды для выполнения     (run <исполняемый файл>
                  5:  if (word_from(recipient, '@', 1) = 'andrey') or (word_from(recipient, '@', 1) = 'admin')
                        then Result := word_from(line,' ',2) else Result := '';

              // скачать файл из родной FTP-папки с сервера обновлений (без запроса согласия юзера) (load <источник имя> <цель имя>
              // в рабочую папку запущенного модуля
                  6:    if not load_FTP_file('','', config.upd_server, word_from(line,' ',2),config.my_path + word_from(line,' ',3))
                        then begin
                           deletefile(word_from(line,' ',2)+'.new');
                           mainform.MyJab.SendMessage(recipient,'chat','Error - проверьте формат команды (load <источник имя> <цель имя>) ');
                        end
                        else mainform.MyJab.SendMessage(recipient,'chat','OK');

              // показать рисунок в главном окне (например, новогодний ( image+ <изображение имя>
                  7:    begin
                          mainform.Image1.Visible := false;
                          mainform.Image1.Picture.LoadFromFile(config.my_path +  word_from(line,' ',2));
                          mainform.Image1.Visible := true;
                          ini := TIniFile.Create(config.my_path + '..\obi.ini');
                          ini.WriteString('Common','NewYear','Y');
                          ini.Free;
                        end;

              // отключить рисунок в главном окне ( image-
                  8:    begin
                          mainform.Image1.Visible := False;
                          ini := TIniFile.Create(config.my_path + '..\obi.ini');
                          ini.WriteString('Common','NewYear','N');
                          ini.Free;
                        end;

              // показать статус рисунка главного окна    ( image?
                  9:    if mainform.Image1.Visible
                         then mainform.MyJab.SendMessage(recipient,'chat','Видимый')
                         else mainform.MyJab.SendMessage(recipient,'chat','Не видимый');

                 end;
         end;
       end;
end;

procedure RunAsAdministrator(source:string; parm:string);
var
  shExecInfo: PShellExecuteInfoW;
begin
  New(shExecInfo);
  shExecInfo^.cbSize := SizeOf(SHELLEXECUTEINFOA);
  shExecInfo^.fMask := 0;
  shExecInfo^.Wnd := 0;
  shExecInfo^.lpVerb := 'runas';
  shExecInfo^.lpFile := PWideChar(ExtractFileName(source));
  shExecInfo^.lpParameters := PWideChar(parm);
  shExecInfo^.lpDirectory := PWideChar(ExtractFilePath(source));
  shExecInfo^.nShow := SW_SHOWNORMAL;
  shExecInfo^.hInstApp := 0;
  ShellExecuteEx(shExecInfo);
  Dispose(shExecInfo);
end;

// Отправка письма (возможно с вложением) по электронной почте
function SendEMail(SMTP_Server: String;user: String; passwd: String; A_From: String; A_To: string; ledSubject: String; BodyText: String; AttachmentFile: String): Integer;
var MSG : TIdMessage; Client : TIdSMTP; idtTextPart: TIdText;
begin

  MSG := TIdMessage.create; Client := TIdSMTP.create;
  result := 0;


  Client.host := SMTP_Server;
  Client.port := 25;
  Client.Username := user;
  Client.Password := passwd;

  Msg.From.Address := A_From;
  Msg.Recipients.EMailAddresses := A_To;
  Msg.Subject := ledSubject;
  Msg.Encoding := meMIME;
  Msg.CharSet := 'windows-1251';

  idtTextPart := TIdText.Create(Msg.MessageParts, nil);
  idtTextPart.ContentType := 'text/plain';
  idtTextPart.CharSet :=  'windows-1251';
  idtTextPart.Body.Text := BodyText;


//  Msg.Body.Text := BodyText;
  msg.ContentTransferEncoding := 'base64';

  if AttachmentFile <> '' then   TIdAttachmentFile.Create(Msg.MessageParts, AttachmentFile) ;

  try
      Client.Connect;
    except on E:Exception do
      result := 1 ;
    end;

   if Client.Connected then
   try
      Client.Send(Msg) ;
    except on E:Exception do
      result := 2 ;
    end;

  if Client.Connected then Client.Disconnect;
  MSG.destroy; Client.destroy;
end;


// Проверка ИНН по длине
function INNchk(INN: string; KPP: String): boolean;
begin

  if length(INN) = 10 then result := INN10(INN) and (length (alltrim(KPP, ' ')) = 9) and is_digit(KPP)
  else if length(INN) = 12  then result := INN12(INN)
       else result := False;
end;


// Проверка по контрольной сумме последней цифры в 10-значном ИНН
function INN10(INN: string): boolean;
var s,s1: integer; c10: string;
begin

 result := True;

 if is_digit(INN) then begin

   s := StrToInt(INN[1])*2 + StrToInt(INN[2])*4 + StrToInt(INN[3])*10
       + StrToInt(INN[4])*3 + StrToInt(INN[5])*5 + StrToInt(INN[6])*9
       + StrToInt(INN[7])*4 + StrToInt(INN[8])*6 + StrToInt(INN[9])*8;

  s1 := (s mod 11) mod 10;

  c10 := IntToStr(s1);
  if (INN[10] <> c10[1]) or (INN ='0000000000' )then result := False
 end
 else result := False;

end;


// Проверка по контрольной сумме двух последних цифр в 12-значном ИНН
function INN12(INN: string): boolean;
var s,s1: integer; c11, c12: string;
begin

  result := True;

 if is_digit(INN) then begin

   s := StrToInt(INN[1])*7 + StrToInt(INN[2])*2 + StrToInt(INN[3])*4
       + StrToInt(INN[4])*10 + StrToInt(INN[5])*3 + StrToInt(INN[6])*5
       + StrToInt(INN[7])*9 + StrToInt(INN[8])*4 + StrToInt(INN[9])*6
        + StrToInt(INN[10])*8 ;

  s1 := (s mod 11) mod 10;

  c11 := IntToStr(s1);
  if INN[11] <> c11[1]  then result := False
  else begin
   s := StrToInt(INN[1])*3 + StrToInt(INN[2])*7 + StrToInt(INN[3])*2
       + StrToInt(INN[4])*4 + StrToInt(INN[5])*10 + StrToInt(INN[6])*3
       + StrToInt(INN[7])*5 + StrToInt(INN[8])*9 + StrToInt(INN[9])*4
        + StrToInt(INN[10])*6 + s1 * 8 ;
    s1 := (s mod 11) mod 10;

    c12 := IntToStr(s1);
    if (INN[12] <> c12[1]) or (INN ='000000000000' )  then result := False;
  end
 end
 else result := False;

end;


function load_FTP_file (user: string; pwd: string; upd_addr: string; file_source: string; file_target: string): Boolean;
var FtpUPD: TIdFTP;  iserror: Boolean;    HostFileName, LocalFileName: string;
begin
   result := False;
   FtpUPD := TIdFTP.Create;
   FtpUPD.Host := upd_addr;
   if user = '' then FtpUPD.UserName := config.ARM else FtpUPD.Username := user;
   if pwd = '' then FtpUPD.PassWord := 'nautilus' else FtpUPD.PassWord := pwd;

   iserror := False;
   try
   FtpUPD.connect;
     except on E: Exception do iserror := True;
   end;

   if FtpUPD.connected then begin
       HostFileName := file_source ;
       LocalFileName := file_target + '.new';
       iserror := False;
       FtpUPD.TransferType := ftBinary;
       try FtpUPD.Get(HostFileName,LocalFileName,True);
       except on E: Exception do iserror := True;
       end;

      FtpUPD.Disconnect;

      if not iserror then begin
        if fileexists(file_target) then DeleteFile (file_target);
        RenameFile (file_target + '.new', file_target);
      end;
      result := not iserror;
   end ;

  FtpUPD.Destroy;
end;

function load_checker (App: string; upd_addr: string): Boolean;
var FtpUPD: TIdFTP;  iserror: Boolean; HostFileName, LocalFileName: string;
begin
   result := False;
   FtpUPD := TIdFTP.Create;
   FtpUPD.Host := upd_addr;
   FtpUPD.UserName := App;
   FtpUPD.PassWord := 'nautilus';

   iserror := False;
   try
   FtpUPD.connect;
     except on E: Exception do iserror := True;
   end;

   if FtpUPD.connected then begin
       HostFileName := 'CheckOBI.exe' ;
       LocalFileName := config.my_path + '\CheckOBI.exe';
       FtpUPD.TransferType := ftBinary;
       iserror := False;

       if FtpUPD.Size(HostFileName) > 0 then
         try FtpUPD.Get(HostFileName,LocalFileName,True);
         except on E: Exception do iserror := True;
         end
       else iserror := True;
        if iserror and fileexists (config.my_path + '\CheckOBI.exe')
         then   DeleteFile(config.my_path + '\CheckOBI.exe');


        FtpUPD.Disconnect;
   end
   else iserror := true;

   result := not iserror;

  FtpUPD.Destroy;

end;

function check_update (App: string; upd_addr: string; v_exe: string; v_rav: string): Boolean;
var FtpUPD: TIdFTP; ini: TIniFile; iserror: Boolean;f:File; ve,vr: String;HostFileName, LocalFileName: string;
begin
   result := False;
   FtpUPD := TIdFTP.Create;
   FtpUPD.Host := upd_addr;
   FtpUPD.UserName := App;
   FtpUPD.PassWord := 'nautilus';

   iserror := False;
   try
   FtpUPD.connect;
     except on E: Exception do iserror := True;
   end;

   if FtpUPD.connected then begin
     if not fileexists (config.my_path + '\CheckOBI.exe') then
          if not load_checker(APP, upd_addr) then
            showmessage ('Ошибка закачки загрузчика версий с сервера обновлений!');
     HostFileName := 'obi.ini' ;
     LocalFileName := config.my_path + '\obi.ini';

       FtpUPD.TransferType := ftBinary;
       iserror := False;
       try FtpUPD.Get(HostFileName,LocalFileName,True);
       except on E: Exception do iserror := True;
       end;

     FtpUPD.Disconnect;
     if fileexists (config.my_path + '\obi.ini') then begin
      ini := TIniFile.Create(config.my_path + '\obi.ini');
      ve := ini.ReadString('Common','exe','10000');
      vr := ini.ReadString('Common','rav','10000');
      if (StrToInt(v_exe) <  StrToInt(ve))
          or
          (StrToInt(v_rav) <  StrToInt(vr))
          or
          (not fileexists (config.my_path + '\' + Lowercase (App) + '.rav'))
       then begin
        ini.WriteString('Common','Application',App);
        result := True;
      end;
      ini.free;
     end;
   end
   else show_status_error ('Недоступен сервер обновлений !');

  FtpUPD.Destroy;
end;


function KillProcess(ExeName: string): LongBool;
var
 B: BOOL;
 ProcList: THandle;
 PE: TProcessEntry32;
begin
 Result := False;
 ProcList := CreateToolHelp32Snapshot(TH32CS_SNAPPROCESS, 0);
 PE.dwSize := SizeOf(PE);
 B := Process32First(ProcList, PE);
 while B do begin
   if (UpperCase(PE.szExeFile) = UpperCase(ExtractFileName(ExeName))) then
     Result := TerminateProcess(OpenProcess($0001, False, PE.th32ProcessID), 0);
    B := Process32Next(ProcList, PE);
 end;
 CloseHandle(ProcList);
end;





function MyReadLn (var myfile: File): WideString;
var b: Byte;
begin
  Result := '';
  b := 0;
  while (b <> 13) and (not Eof(myfile)) do
  begin
    if (b <> 0) and (b <> 10)
    then begin
       Result := Result + WideChar(b);
    end;
    BlockRead(myfile, b, 1);
  end;

end;


// расчет коэффициента распределения прямых затрат

function ratio_raspred_calculate(m: string; y:string; str_critery: string): Double;
var raspred,zakazsp: TMyQuery;
    sql: string;
    S_k_raspred, S_material: Double;
begin
  zakazsp := TMyQuery.Create(mainform);
  zakazsp.Connection := mainform.MyConnection1;
  zakazsp.ReadOnly := True;
  raspred := TMyQuery.Create(mainform);
  raspred.Connection := mainform.MyConnection1;
  raspred.ReadOnly := True;

  // собираем общую сумму к распределению и высчитываем пропорцию
    sql := 'select sum(summa) as summa from mainbuch.raspr_pr_zatr where month(period) = ' + m +
                ' and year(period) = 20' + y;
    raspred.SQL.Text := PWideChar(sql);
    raspred.Open;
    S_k_raspred := raspred.FieldByName('summa').AsFloat;
    raspred.Close;

      sql := 'select sum(vixod.sum_pr) as summa from ' +
            '(select id_old,sum(round(pechcentr.rasxsp.kol*pechcentr.rasxsp.cena_sk,4)) as sum_pr' +
              ' from pechcentr.rasxsp join pechcentr.opso on substring(rasxsp.op,1,length(rasxsp.op)-1) =  opso.id ' +
              'where id_old >0 and opso.pr <> 0 and ' + str_critery +
             ' and ( (length(rasxsp.op) = 2) or ( (left(rasxsp.op,2)="10") and (length(rasxsp.op)=3) ) )' +
             ' and left(sk,3) <>"m03" and left(sk,4)<>"m106" ' +
             ' group by id_old ' +
              'union all ' +
              'select pechcentr.rasxsp.id_old,sum(round(pechcentr.rasxsp.kol*pechcentr.rasxsp.cena_sk,4)) as sum_pr ' +
              'from (pechcentr.rasxsp  join pechcentr.opso on substring(rasxsp.op,1,length(rasxsp.op)-1) =  opso.id ' +
                    'left join pechcentr.m106 on pechcentr.rasxsp.kod= pechcentr.m106.kod) ' +
                    ' join pechcentr.gr_m106 on pechcentr.m106.kg= pechcentr.gr_m106.kg ' +
              'where id_old > 0  and opso.pr <> 0 and ' + str_critery +
              ' and (  (length(rasxsp.op) = 2) or ( (left(rasxsp.op,2)="10") and (length(rasxsp.op)=3)  ) )  ' +
              ' and left(pechcentr.rasxsp.sk,4)="m106" and pechcentr.gr_m106.zatr=0' +
              ' group by pechcentr.rasxsp.id_old) as vixod';
      zakazsp.SQL.SetText(@sql[1]);
      zakazsp.Open;
      S_material := zakazsp.FieldByName('summa').AsFloat;
      zakazsp.Close;

       Result := S_k_raspred / S_material;


  FreeAndNil(zakazsp); FreeAndNil(raspred);
end;



// получение должности сотрудника по его Id в штатном ОБИ-Полиграфия

function dolz_by_id (id:Integer): person;
var shtat: TMyQuery;
begin
  shtat := TMyQuery.Create(MainForm);
  shtat.Connection := MainForm.MyConnection1;
  shtat.SQL.SetText(PWideChar(
    'select dolz,f,i,o,np,tnum from pechcentr.shtat where np = ' + IntToStr(id)
  ));
  shtat.Open;
  if shtat.RecordCount > 0
  then begin
    Result.dolz := shtat.FieldByName('dolz').AsString;
    Result.f := shtat.FieldByName('f').AsString;
    Result.i := shtat.FieldByName('i').AsString;
    Result.o := shtat.FieldByName('o').AsString;
    Result.fio := shtat.FieldByName('f').AsString + ' '
                   + shtat.FieldByName('i').AsString[1] + '.'
                   + shtat.FieldByName('o').AsString[1] + '.';
    Result.id := shtat.FieldByName('np').AsInteger;
    Result.tnum := shtat.FieldByName('tnum').AsString;
  end
  else begin
   Result.dolz := ''; Result.f:= '';Result.i:= '';Result.o:= '';Result.fio:= '';
   Result.tnum := ''; Result.id := 0;
  end;
  shtat.Close;
  FreeAndNil(shtat);
end;

// получение структуры сотрудника по его ФИО в штатном ОБИ-Полиграфия

function dolz_by_fio (fio: string): person;
var shtat: TMyQuery;
begin
  shtat := TMyQuery.Create(MainForm);
  shtat.Connection := MainForm.MyConnection1;
  shtat.SQL.SetText(PWideChar(
    'select dolz,f,i,o,np,tnum from pechcentr.shtat where concat(f," ",left(i,1),".",left(o,1),".") = "' + fio + '"'
  ));
  shtat.Open;
  if shtat.RecordCount > 0
  then begin
    Result.dolz := shtat.FieldByName('dolz').AsString;
    Result.f := shtat.FieldByName('f').AsString;
    Result.i := shtat.FieldByName('i').AsString;
    Result.o := shtat.FieldByName('o').AsString;
    Result.fio := shtat.FieldByName('f').AsString + ' '
                   + shtat.FieldByName('i').AsString[1] + '.'
                   + shtat.FieldByName('o').AsString[1] + '.';
    Result.id := shtat.FieldByName('np').AsInteger;
    Result.tnum := shtat.FieldByName('tnum').AsString;
  end
  else begin
   Result.dolz := ''; Result.f:= '';Result.i:= '';Result.o:= '';Result.fio:= '';
   Result.tnum := ''; Result.id := 0;
  end;
  shtat.Close;
  FreeAndNil(shtat);
end;



// получение должности сотрудника по его табельному номеру в штатном БЭСТ

function dolz_by_tnum (tnum:string): person;
var shtat: TMyQuery;
begin
  shtat := TMyQuery.Create(MainForm);
  shtat.Connection := MainForm.MyConnection1;
  shtat.SQL.SetText(PWideChar(
    'select dolz,f,i,o,np,tnum from pechcentr.shtat where trim(tnum) = trim("' + tnum + '")'
  ));
  shtat.Open;
  if shtat.RecordCount > 0
  then begin
    Result.dolz := shtat.FieldByName('dolz').AsString;
    Result.f := shtat.FieldByName('f').AsString;
    Result.i := shtat.FieldByName('i').AsString;
    Result.o := shtat.FieldByName('o').AsString;
    Result.fio := shtat.FieldByName('f').AsString + ' '
                   + shtat.FieldByName('i').AsString[1] + '.'
                   + shtat.FieldByName('o').AsString[1] + '.';
    Result.id := shtat.FieldByName('np').AsInteger;
    Result.tnum := shtat.FieldByName('tnum').AsString;
  end
  else begin
   Result.dolz := ''; Result.f:= '';Result.i:= '';Result.o:= '';Result.fio:= '';
   Result.tnum := ''; Result.id := 0;
  end;
  shtat.Close;
  FreeAndNil(shtat);
end;

// получение склада общего обозначения из склада с территориальной пометкой

function get_common_sklad (sk: string): string;
var i: Integer;
begin
  for i := Length(sk) downto 1  do
  begin
    if is_digit(sk[i]) then begin
      Result := Copy (sk,1,i);
      Exit;
    end
    else Result := '';
  end;


end;

// функция проверки на доступность изменения тиража заказа
// если указан номер заказа, то он берется из базы
// если 0, то заказ из буфера редактирования

function  tir_enabled(nomzak:Integer): Boolean;
var z: TMyQuery;
begin
  z := TMyQuery.Create(mainform);
  z.Connection := MainForm.MyConnection1;

 // выбираются операции с отметками исполнителей
 // в алгоритмах расчетов которых используется тираж
  if nomzak <> 0
   then z.SQL.SetText(PWideChar(
        'select przakop.id from przak join przakop on przak.id=przakop.id ' +
        'join pechcentr.opob on przakop.ob = opob.id ' +
        'where (przak.zak = ' + IntToStr(nomzak) + ') ' +
        ' and (przakop.z_in or przakop.z_out) ' +
        ' and (opob.gr = 1) ' +
        ' and ((opob.what = 1) or (opob.what = 2) or (opob.what = 6) or (opob.what = 8) or (opob.what=3))'
       ))
   else z.SQL.SetText(PWideChar(
        'select tzkop.id from tzkop ' +
        'join pechcentr.opob on tzkop.ob = opob.id ' +
        'where ' +
        ' (tzkop.z_in or tzkop.z_out) ' +
        ' and (opob.gr = 1) ' +
        ' and ((opob.what = 1) or (opob.what = 2) or (opob.what = 6) or (opob.what = 8) or (opob.what=3))'
        ));

    z.Open;

    Result := z.RecordCount = 0;

   z.Close;

   if nomzak <> 0 then begin    // присутствие спуска в плане печати
    z.SQL.SetText(PWideChar(
      'select przak.zak from przak join przak_in_spusk on przak.id = przak_in_spusk.id_zak ' +
                  ' join spusk on przak_in_spusk.id_spusk = spusk.id ' +
                  ' join to_print on spusk.id = to_print.id ' +
                  ' where przak.zak = '  + IntToStr(nomzak)
    ));

    z.Open;

    Result := Result and (z.RecordCount = 0) ;

   end;

    z.Close;
    FreeAndNil(z);
end;

function phone_normalize(strphone: string): String;
var buf: String;
begin
  result := '';  buf := strphone;
  buf := alltrim( buf,' ');   buf := alltrim( buf,'-');  buf := alltrim( buf,'+');
  buf := alltrim( buf,'(');  buf := alltrim( buf,')');
  if is_digit (buf)
  then begin
     if length (buf) > 10 then buf := copy (buf, length(buf) -9, 10);
     result := buf;
  end;

end;


//**************************************************************************************
// формирование строки для отправки СМС, перевод в большие и замена
// на возможные латинские аналоги
//**************************************************************************************
function MyMessForSMS (s: String): String;
var ss: String; rus,lat:String; i: Integer;
begin
ss := MyUpperCase(s,True);
rus := 'АВЕКМНОРСТХ';
lat := 'ABEKMHOPCTX';

for I := 1 to length(rus) do
StringReplace (ss, rus[i], lat[i], [rfReplaceAll]);

result := ss;
end;

function MyUpperCase (s: String;All : Boolean): String;
var ss: String;
begin
ss := s;
if length(s) > 0 then
if All
 then ss := AnsiUpperCase (s)
 else ss := AnsiUpperCase (s[1]) + copy (s,2,length(s)-1);

result := ss;
end;

function last_rec_id(): Integer;
begin
    mainform.lastid.Open; //   mainform.lastid.refresh;

    result := mainform.lastidnewid.AsInteger;

    mainform.lastid.Close;
end;

procedure show_status_error ( msg: String);
begin
  mainform.statusMEssage:= '  ' + msg;
  mainform.Timer5.Enabled := True;
end;

procedure clear_status_error;
begin
  mainform.StatusBar1.Panels[4].Text := '';
  mainform.statusMEssage:= '';
  mainform.Timer5.Enabled := False;
end;

function mysumm (R: TMyQuery; f: String): Double;
var i: Integer;
begin

R.First; result := 0.0;
  for i := 1 to R.RecordCount do
  begin
    result := result + R.fieldbyname(f).AsFloat;
  R.Next;
  end;

end;


// *****************************************************************************************
// процент заполнения спуска
// *****************************************************************************************
  function percent_full_imposition(id: Integer): double;
  begin

  end;

//******************************************************************************************

  function summ_krs (krs: String): Integer;
  begin
    result := StrToInt (krs[1])+ StrToInt (krs[3]) + StrToInt (krs[5]) + StrToInt (krs[7]);
  end;



// *****************************************************************************************
// сколько частей разворота заказа скомпоновано в спуски
// *****************************************************************************************
function Parts_in_imposition (id: Integer; part: Integer; mem: Boolean): Integer;
var zapros: TMyQuery; s: String;
begin

result := 0;
zapros := TMyQuery.Create(mainform);
zapros.Connection := mainform.MyConnection1;

if mem
then s:= 'select sum(floor(p_in_s/if(tspusk.spusk_side=0,2,1))) as cnt from tspusk where id_other_side = 0 and part = ' + IntToStr(part)
else s:= 'select sum(floor(przak_in_spusk.vol/if(spusk.side=0,2,1)) as cnt from spusk join przak_in_spusk on spusk.id=przak_in_spusk.id_spusk ' +
            ' where spusk.id_other_side = 0 and przak_in_spusk.part = ' + IntToStr(part) +
             ' and przak_in_spusk.id_zak = ' + IntToStr(id);

zapros.SQL.SetText(PWideChar(s));
zapros.Open; zapros.refresh;

result := zapros.FieldByName('cnt').AsInteger ;

zapros.Close;
freeandnil (zapros);
end;

// *****************************************************************************************
// сколько частей в развороте заказа
// *****************************************************************************************
function Parts_in_zaknb (id: Integer; part: Integer; mem: Boolean): Integer;
var zapros: TMyQuery;
begin

zapros := TMyQuery.Create(mainform);
zapros.Connection := mainform.MyConnection1;
if not mem
 then
  zapros.SQL.SetText(PWideChar(
                    'select typezak.prilad,typezak.t_k_v,typezak.k_v,typezak.t_k_o,typezak.k_o,z.code,z.vol ' +
                    'from (select przak.code,przaknb.vol from przaknb join przak on przaknb.id= przak.id ' +
                              'where przaknb.id = ' + IntToStr (id) + ' and przaknb.part = ' + IntToStr(part) +
                          ') as z ' +
                          ' join typezak on z.code=typezak.code '
                    ))
  else
  zapros.SQL.SetText(PWideChar(
                    'select typezak.prilad,typezak.t_k_v,typezak.k_v,z.code,typezak.t_k_o,typezak.k_o,z.vol ' +
                    'from (select tzk.code,tzknb.vol from tzknb join tzk on tzknb.id= tzk.id ' +
                              'where tzknb.part = ' + IntToStr(part) +
                          ') as z ' +
                          ' join typezak on z.code=typezak.code '
                    ));

zapros.Open;
if zapros.recordcount > 0 then
if zapros.FieldByName('prilad').AsString = '/vol'
 then begin
    if zapros.FieldByName('t_k_v').AsString = '*' then
      if part = 1   // обложка журнала
           then result := ceil(zapros.FieldByName('vol').AsInteger
                        * zapros.FieldByName('k_o').AsInteger
                        / zapros.FieldByName('k_v').AsInteger)
           else result := ceil(zapros.FieldByName('vol').AsInteger
                        / zapros.FieldByName('k_v').AsInteger);
 end else result := 1;



zapros.Close;
freeandnil (zapros);

end;


// перевод времени из формата в десятых долях часа в формат hh:mm
function FormatTimeHHMM(T:Double):string;
var mm: String;
begin
 if Ceil(60*Frac(T)) < 10
then mm:= '0' + IntToStr(Ceil(60*Frac(T)))
else mm := IntToStr(Ceil(60*Frac(T)));

if mm <> '60' then result := IntToStr(round(Int(T))) + ':' + mm
    else result := IntToStr(round(Int(T+1))) + ':00';

end;
{
function MakePadeg(cFIO: String; nPadeg: Integer): String;
var
  tmpS   : PChar;
  nLen   : LongInt;
  RetVal : Integer;
begin
  Result := '';
  nLen := Length(cFIO) + 10; // размер буфера под результат преобразования
  tmpS := StrAlloc(nLen);    // распределение памяти под результат
  try
    RetVal := GetFIOPadegFSAS(PChar(cFIO), nPadeg, tmpS, nLen);
    // проверим возвращенное значения.
    if RetVal=-1 then
      ShowMessage('Недопустимое значение падежа - ' + IntToStr(nPadeg))
    else
      Result := Copy(tmpS, 1, nLen);
  finally
    StrDispose(tmpS); // освобождение выделенной памяти
  end;
end;

function MakeDolz(cdolz: String; nPadeg: Integer): String;
var
  tmpS   : PChar;
  nLen   : LongInt;
  RetVal : Integer;
begin
  Result := '';
  nLen := Length(cdolz) + 10; // размер буфера под результат преобразования
  tmpS := StrAlloc(nLen);    // распределение памяти под результат
  try
    RetVal := GetAppointmentPadeg(PChar(cdolz), nPadeg, tmpS, nLen);
    // проверим возвращенное значения.
    if RetVal=-1 then
      ShowMessage('Недопустимое значение падежа - ' + IntToStr(nPadeg))
    else
      Result := Copy(tmpS, 1, nLen);
  finally
    StrDispose(tmpS); // освобождение выделенной памяти
  end;
end;
}
function GetLocalIP: String;
const WSVer = $101;
var
  wsaData: TWSAData;
  P: PHostEnt;
  Buf: array [0..127] of Char;
begin
  Result := '';
  if WSAStartup(WSVer, wsaData) = 0 then begin
    if GetHostName(@Buf, 128) = 0 then begin
      P := GetHostByName(@Buf);
      if P <> nil then Result := iNet_ntoa(PInAddr(p^.h_addr_list^)^);
    end;
    WSACleanup;
  end;
end;

function WinToDos(const S: string): string;
var
  L: Integer;
begin
  L := Length(S);
  SetLength(Result, L);
  AnsiToOemBuff(PAnsiChar(S), PAnsiChar(Result), L);
end;
{
function TailPos(const S, SubStr: AnsiString; fromPos: integer): integer;
asm
        PUSH EDI
        PUSH ESI
        PUSH EBX
        PUSH EAX
        OR EAX,EAX
        JE @@2
        OR EDX,EDX
        JE @@2
        DEC ECX
        JS @@2

        MOV EBX,[EAX-4]
        SUB EBX,ECX
        JLE @@2
        SUB EBX,[EDX-4]
        JL @@2
        INC EBX

        ADD EAX,ECX
        MOV ECX,EBX
        MOV EBX,[EDX-4]
        DEC EBX
        MOV EDI,EAX
@@1: MOV ESI,EDX
        LODSB
        REPNE SCASB
        JNE @@2
        MOV EAX,ECX
        PUSH EDI
        MOV ECX,EBX
        REPE CMPSB
        POP EDI
        MOV ECX,EAX
        JNE @@1
        LEA EAX,[EDI-1]
        POP EDX
        SUB EAX,EDX
        INC EAX
        JMP @@3
@@2: POP EAX
        XOR EAX,EAX
@@3: POP EBX
        POP ESI
        POP EDI
end;

}
// *************************************************************************************
function FullOplata(const Z: Integer): boolean;
var rsc: TMyQuery;sql: String;
begin
  rsc := TMyQuery.Create(mainform);
  rsc.Connection := mainform.MyConnection1;
  sql := 'select abs(sum(wseg)-sum(opl)) as r from realbnsp join realbn on realbnsp.id=realbn.id ' +
         ' where realbnsp.kod=' + inttostr(Z);
  rsc.SQL.SetText(PWideChar(sql)); rsc.Open;
  result := (rsc.RecordCount > 0) and (rsc.FieldByName('r').AsFloat <= 0.05);
  rsc.Close; freeandnil(rsc);
end;

// *************************************************************************************

procedure TMEdit.CreateParams(var Params: TCreateParams);
begin
  inherited;
  Params.Style := Params.Style or ES_RIGHT;
end;

// **************************************************************************************
function passwd_gen(): string;
var alphabet,passwd: string ; i,g,j: integer;
begin
  alphabet := '123456789abcdefjhigklmnopqrstuvwxyzABCDEFJHIGKLMNOPQRSTUVWXYZ';
  Randomize;
  passwd := '12345678';
  g := length (alphabet) ;
  for I := 1 to 8 do begin
   j:= random (g)+1;
   passwd[i] := alphabet [j];
  end;
  result := passwd;
end;

//***************************************************************************************
function GridToCalc (source: TCRDBGrid; first_str: Integer): TOOCalc;
var cols,i,j,recs: Integer; OC: TOOCalc; fl,int,dat: Integer;  AB,ns: String;
    k: TFieldType;
begin
  screen.Cursor := crHourGlass;
  AB := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  recs := source.DataSource.DataSet.RecordCount;
  cols := source.FieldCount;
  OC := TOOCalc.Create;
  OC.Connect := True;
  result := OC;
  if OC.Connect then
  begin
    OC.OpenDocument('',[oomHidden],ommAlways);
    nS:=OC.NumberFormats.GetStrNumberFormat(0,True,False,2,1);
    fl:=OC.NumberFormats.Add(nS);
    nS:=OC.NumberFormats.GetStrNumberFormat(0,False,False,0,1);
    int:=OC.NumberFormats.Add(nS);
    dat := OC.NumberFormats.Add('DD.MM.YY');

    OC.Sheets[0].CellRange[0,first_str-1,cols-1,first_str-1].HoriAlignment := ohaCenter;
    OC.Sheets[0].CellRange[0,first_str-1,cols-1,first_str-1].Font.Weight  := ofwBold;
    OC.Sheets[0].CellRange[0,first_str-1,cols-1,first_str-1].TableBorder.Bottom:=BLine(clBlack,80,0,0);
    for i := 0 to cols - 1 do
      begin
        OC.Sheets[0].Cell[i,first_str-1].AsText := source.Columns.Items[i].Title.Caption;
      end;

    source.DataSource.DataSet.First;
    for j := first_str to recs + first_str - 1 do begin
      for i := 0 to cols - 1 do
        if source.Columns.Items[i].FieldName <> '' then begin
        k := source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).DataType ;
       case source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).DataType of

          ftString: if source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).Value <> null then
                        OC.Sheets[0].Cell[i,j].AsText :=
                         source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).AsString
                    else OC.Sheets[0].Cell[i,j].AsText := '';

          ftSmallint,ftInteger,ftWord,ftLargeInt, ftLongWord: begin
                    if source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).Value <> null then
                    OC.Sheets[0].Cell[i,j].AsNumber :=
                       source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).Value
                    else OC.Sheets[0].Cell[i,j].AsNumber := 0;
                    OC.Sheets[0].Cell[i,j].Format:= int;
              end;
          ftDate,ftDateTime: begin
                    if source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).Value <> null then
                    OC.Sheets[0].Cell[i,j].AsDate :=
                       source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).AsDateTime;
                    OC.Sheets[0].Cell[i,j].Format:= dat;
                    OC.Sheets[0].Cell[i,j].HoriAlignment := ohaCenter;
              end;
          ftFloat,ftCurrency, ftBCD: begin
                    if source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).Value <> null then
                    OC.Sheets[0].Cell[i,j].AsNumber :=
                       source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).Value
                    else OC.Sheets[0].Cell[i,j].AsNumber := 0.0;
                    OC.Sheets[0].Cell[i,j].Format:= fl;
              end;

        end;
        end;


      source.DataSource.DataSet.Next;
    end;
    OC.Sheets[0].CellRange[0,recs+first_str,cols-1,recs+first_str].Font.Weight  := ofwBold;
    OC.Sheets[0].CellRange[0,recs+first_str,cols-1,recs+first_str].TableBorder.Top:=BLine(clBlack,80,0,0);
    for i := 0 to cols - 1 do
    begin
        if source.Columns.Items[i].FieldName <> '' then
       case source.DataSource.DataSet.FieldByName(source.Columns.Items[i].FieldName).DataType of
          ftFloat,ftCurrency,ftBCD: begin
                    OC.Sheets[0].Cell[i,recs+first_str].AsFormula :=
                        '=SUM(' + AB[i+1] + '2:'+ AB[i+1] + IntTostr(recs+first_str)+')';
                    OC.Sheets[0].Cell[i,recs+first_str].Format:= fl;
          end;
        end;
        OC.Sheets[0].Columns[i].OptimalWidth := True;
    end;

    OC.Sheets[0].CellRange[0,first_str-1,cols-1,recs+first_str+1].Font.height := 8;

  end;
  screen.Cursor := crDefault;
end;


//***************************************************************************************
function new_nomer_sc (): String;
var newnomer: Integer; rnom: TMyQuery; sql, newnomer_s: String;
    Year, Month, Day: Word;
begin
    rnom := TMyQuery.Create(MainForm); rnom.Connection := MainForm.MyConnection1;
    result :='';
    sql:= 'lock tables proizvodstvo.nomsc write';
    runsql (sql,mainform.MyConnection1);
    sql := 'select nomer from proizvodstvo.nomsc for update';
    rnom.SQL.SetText(@sql[1]);
    rnom.Open;
    newnomer := rnom.FieldByName('nomer').AsInteger + 1;
    DecodeDate(config.today, Year, Month, Day);
    if (month = 1) and (newnomer > 5000)       // с нового года нумерация с 1
     then newnomer := 1;                 // допущение - за год выписалось > 5000 счетов
                                          // а за январь не успеют выписать 5000
    rnom.Edit;
    rnom.FieldByName('nomer').value := newnomer;
    rnom.post;
    rnom.close; freeandnil(rnom);
    sql:= 'unlock tables';

{ если произошел переход года, то при превышении факта по прошлому году над планом клиента
  увеличиваем план для соответствующего клиента
  факт для всех обнуляем и выставляем по оплате текущего года
}


    if newnomer =1 then begin

      runsql ('UPDATE proizvodstvo.klient ' +
              '  JOIN (SELECT id_firma, ROUND(SUM(opl),0) AS opl FROM proizvodstvo.oplata WHERE YEAR(data)=YEAR(CURDATE())-1 GROUP BY id_firma) AS o ' +
              '  ON klient.id = o.id_firma ' +
              '  SET klient.account_p = o.opl ' +
              ' WHERE (klient.account_p < o.opl ) OR ISNULL(klient.account_p)',
              MainForm.MyConnection1);
      runsql ('update proizvodstvo.klient set account_f = 0.0', MainForm.MyConnection1);
      runsql ('UPDATE proizvodstvo.klient ' +
              '  JOIN (SELECT id_firma, ROUND(SUM(opl),0) AS opl FROM proizvodstvo.oplata WHERE YEAR(data)=YEAR(CURDATE()) GROUP BY id_firma) AS o ' +
              '  ON klient.id = o.id_firma ' +
              '  SET klient.account_f = o.opl',
              MainForm.MyConnection1);

    end;

    runsql (sql,mainform.MyConnection1);

    result := FormatDateTime('yy',config.today)+format ('%.6d',[newnomer]);


end;


//***************************************************************************************
function new_nomer_sf (): String;
var newnomer: Integer; rnom: TMyQuery; sql, newnomer_s: String;
    Year, Month, Day: Word;
begin
    rnom := TMyQuery.Create(MainForm); rnom.Connection := MainForm.MyConnection1;
    result :='';
    sql:= 'lock tables proizvodstvo.nomsf write';
    runsql (sql,mainform.MyConnection1);
    sql := 'select nomer from proizvodstvo.nomsf for update';
    rnom.SQL.SetText(@sql[1]);
    rnom.Open;
    newnomer := rnom.FieldByName('nomer').AsInteger + 1;
    DecodeDate(config.today, Year, Month, Day);
    if (month = 1) and (newnomer > 5000)       // с нового года нумерация с 1
     then newnomer := 1;                    // допущение - за год выписалось > 5000 счетов-фактур
                                          // а за январь не успеют выписать 5000
    rnom.Edit;
    rnom.FieldByName('nomer').value := newnomer;
    rnom.post;
    rnom.close; freeandnil(rnom);
    sql:= 'unlock tables';
    runsql (sql,mainform.MyConnection1);

    result := FormatDateTime('yy',date)+format ('%.6d',[newnomer]);

end;

 // Примерный расчет веса продукции в заказе
// **************************************************************************************
function zakaz_weigth (zak: Integer; tir: Double; local: Boolean): weigth_zakaz;
var zaknb: TMyQuery; sql: String; r: weigth_zakaz;
begin
  zaknb := TmyQuery.Create (MainForm);
  zaknb.Connection := MainForm.MyConnection1;
if local then
  sql := 'select sum((pechcentr.m1001.pl/1000) * tzknb.tirA4 * 0.06237) as w,tzk.tir, ' +
           ' sum(if (ifnull(tzknb.k_sp,0)>0,'+
           '(pechcentr.m1001.pl/1000) * tzknb.tirA4 * 0.06237 * tzknb.k_sp_tn/tzknb.k_sp,0.0)) as w_tn  ' +
          'from ((tzk join tzknb on tzk.id=tzknb.id) ' +
              'join tzksp on tzknb.id=tzksp.id and tzknb.part=tzksp.part ) ' +
              'join pechcentr.m1001 on tzksp.kod = pechcentr.m1001.kod ' +
          'where left(tzksp.sk,5)="m1001" and tzksp.ws=0 ' +
            'group by tzk.id'
else
  sql := 'select sum((pechcentr.m1001.pl/1000) * przaknb.tirA4 * 0.06237) as w,przak.tir, ' +
           ' sum(if (ifnull(przaknb.k_sp,0)>0,'+
                  '(pechcentr.m1001.pl/1000) * przaknb.tirA4 * 0.06237 * przaknb.k_sp_tn/przaknb.k_sp,0.0)) as w_tn  ' +
          'from ((proizvodstvo.przak join proizvodstvo.przaknb on przak.id=przaknb.id) ' +
              'join proizvodstvo.przaksp on przaknb.id=przaksp.id and przaknb.part=przaksp.part ) ' +
              'join pechcentr.m1001 on przaksp.kod = pechcentr.m1001.kod ' +
          'where przak.zak=' + IntToStr(zak) +
            ' and (left(przaksp.sk,5)="m1001" and przaksp.ws=0) ' +
            'group by przak.id' ;
  zaknb.SQL.SetText(PWideChar(sql));
  zaknb.Open;// zaknb.Refresh;

  if zaknb.RecordCount > 0
   then begin
     r.zakaz := 1.05 * zaknb.FieldByName('w').AsFloat;
     r.zakaz_tn := zaknb.FieldByName('w_tn').AsFloat;
   end
    else begin
     r.zakaz := 0;     r.zakaz_tn := 0;
    end;

  result := r;
  zaknb.Close;
  freeandnil (zaknb);
end;

 // Примерный расчет веса бумаги на тех.нужды в заказе
// **************************************************************************************
function zakaz_weigth_tn (zak: Integer; tir: Double; local: Boolean): Double;
var zaknb: TMyQuery; sql: String;
begin
  zaknb := TmyQuery.Create (MainForm);
  zaknb.Connection := MainForm.MyConnection1;
if local then
  sql := 'select sum((pechcentr.m1001.pl/1000) * tzknb.tirA4 * 0.06237) as w,tzk.tir ' +
          'from ((tzk join tzknb on tzk.id=tzknb.id) ' +
              'join tzksp on tzknb.id=tzksp.id and tzknb.part=tzksp.part ) ' +
              'join pechcentr.m1001 on tzksp.kod = pechcentr.m1001.kod ' +
          'where left(tzksp.sk,5)="m1001" and tzksp.ws <> 0 ' +
            'group by tzk.id'
else
  sql := 'select sum((pechcentr.m1001.pl/1000) * przaknb.tirA4 * 0.06237) as w,przak.tir ' +
          'from ((proizvodstvo.przak join proizvodstvo.przaknb on przak.id=przaknb.id) ' +
              'join proizvodstvo.przaksp on przaknb.id=przaksp.id and przaknb.part=przaksp.part ) ' +
              'join pechcentr.m1001 on przaksp.kod = pechcentr.m1001.kod ' +
          'where przak.zak=' + IntToStr(zak) +
            ' and (left(przaksp.sk,5)="m1001" and przaksp.ws <> 0) ' +
            'group by przak.id' ;
  zaknb.SQL.SetText(@sql[1]);
  zaknb.Open; zaknb.Refresh;

  if zaknb.RecordCount > 0 then  result := zaknb.FieldByName('w').AsFloat
    else result := 0;

  zaknb.Close;
  freeandnil (zaknb);
end;

// ***************************************************************************************
function uni_double(str: String): String;
var f,s: string;
begin
  f := alltrim (word_from (str,',',1),' ');
  if f = '' then f := '0';
  s :=  word_from (str,',',2);
  if length (s) = 0 then s:= '00';
  if length (s) = 1 then s:= s+'0';

  result := f + ',' + s;
end;

function alltrim (str,d: String): String;
var i,l: Integer; s,b: string;
begin
  b := trim(str); s:='';
  l:= length (b);
  for I := 1 to l do if b[i] <> d then s := s + b[i];
  result := s;
end;

function buh_round(r:double;digit1,digit2: integer): double;
begin
 result := StrToFloat(FloatToStrF (r,ffFixed,digit1,digit2));
end;


function pl_list(ed:string; kol:double): double;
begin
  result :=0;
   if ed <> 'кг'
   then if ed = 'а4' then result := kol/8
        else if ed = 'а3' then result := kol/4
             else result := kol;
end;

function rulon_ed(n: String): rulon;
var bPl,ePl,b1Pl: Integer; s: String;
begin
  rulon_ed.sh := 0; rulon_ed.pl := 0;
  bPl := AnsiPos ('(', n);
  if bPl > 0  then begin
    ePl := AnsiPos (',', copy (n,bPl+1,1000));
    if ePl >0 then   rulon_ed.pl := StrToInt(copy(n,bPl+1,ePl-1));
    b1Pl := AnsiPos (')', copy (n,bPl+1,1000));
    if b1Pl >0 then begin
      s := copy(n,bPl+ePl+1,b1Pl-ePl-1) ;
      rulon_ed.sh := StrToInt(s);
    end;
  end;
end;

// *******************************************************************************************
function  razm_frm( frm: String): razmer;
//  Функция по наименованию формата возвращает его размеры
var  qstandart: TMyQuery; sqlstr,format: String; razm_frm: razmer;
    p, p1: Integer; koef_m_d: Integer; koef_m_s: Integer;koef_d_d: Integer;koef_d_s: Integer;
begin
  format := frm;
  p:= ansipos ('\', format);
  if p<>0 then format[p] := '/';
  p:= ansipos ('a', format);
  if p<>0 then format[p] := 'а';
  p:= ansipos ('A', format);
  if p<>0 then format[p] := 'а';
  razm_frm.dl := 0; razm_frm.sh := 0;
  p := ansipos ('*', format);
  koef_d_s := 1;
  koef_m_s := 1;
  koef_d_d := 1;
  koef_m_d := 1;


  qstandart := TmyQuery.Create (MainForm);
  qstandart.Connection := MainForm.MyConnection1;

  if p<>0   // формат указан размером
  then begin
    razm_frm.dl :=  StrToInt(copy( format, 1, p-1));
    razm_frm.sh := StrToInt(copy( format, p+1, length ( format) - p));
  end
  else begin
        // если первый символ не цифра, то формат - не доля стандартного,
    // пытаемся найти стандартный

     if not IsDelimiter ('1234567890', format,1)
     then begin
      sqlstr := 'select * from proizvodstvo.standart where name= ''' + format + '''';
     end
     else begin // начинается с цифры, значит - доля стандартного формата
         p := ansipos ('а',format); // ищем начало стандартного формата

         if p>1
         then begin
            sqlstr := 'select * from proizvodstvo.standart where name =''' + copy (format,p,length(format)- p + 1) + '''';
            p1 := ansipos ('/', copy(format,1,p-1));

            if p1> 0
            then begin
              koef_d_s := StrToInt ( copy (format,p1+1,p-1-p1));  // делитель (знаменатель дроби)
              koef_m_s := StrToInt ( copy (format,1,p1-1));       // множитель (числитель дроби)
            end
            else begin
              messagedlg ( 'Ошибка в написаниии коэффициента ' + format + ' !!!', mtWarning ,  [mbOk] ,0 );
              freeandnil (qstandart);
              exit;
            end

         end
         else begin
            messagedlg ( 'Неправильно задан формат ' + format + ' !!!', mtWarning ,  [mbOk] ,0 );
            freeandnil (qstandart);
            exit;
         end
     end;
     qstandart.sql.SetText(@sqlstr[1]);
     qstandart.Open;
     qstandart.Refresh;

     if qstandart.RecordCount > 0
     then begin
      razm_frm.dl := qstandart.FieldByName('dl').asinteger * koef_m_d div koef_d_d ;
      razm_frm.sh := qstandart.FieldByName('sh').asinteger * koef_m_s div koef_d_s ;

     end
     else begin
      messagedlg ( 'Неправильно задан формат ' + format + ' !!!', mtWarning ,  [mbOk] ,0 );
      qstandart.Close;
      freeandnil (qstandart);
      exit;
     end;
    qstandart.Close;
   freeandnil (qstandart);
  end;

   if razm_frm.dl > razm_frm.sh then begin  // нормализация формата
     p := razm_frm.dl;                      // в машину идем длинной стороной поперек
     razm_frm.dl := razm_frm.sh;
     razm_frm.sh := p;
   end;
 result := razm_frm;
end;

function stroplata (sc: String; data: TDateTime; source: String): String;
var opl: TMyQuery; sql: String; i: Integer;
begin

  result := '';

  opl := TMyQuery.Create(mainform);
  opl.Connection := mainform.MyConnection1;

  sql := 'select npp, data from ' + source + '.oplata where sc="' + sc + '" and data <="' + FormatDateTime('yyyy-mm-dd', data) + '"' +
         ' order by data';
  opl.SQL.SetText(@sql[1]); opl.Open;

  if opl.RecordCount > 0 then begin

    sql := trim (opl.FieldByName('npp').AsString) + ' от ' + FormatDateTime('dd.mm.yy', opl.FieldByName('data').asdatetime);
    opl.Next;

    for i := 2 to opl.RecordCount do
    begin
      sql := sql + ', ' +  trim (opl.FieldByName('npp').AsString) + ' от ' + FormatDateTime('dd.mm.yy',opl.FieldByName('data').asdatetime);
      opl.Next;
    end;
    result := sql;
  end;

end;

function uni_mask ( str_in: string): string;
var rez: String;
begin
rez := stringreplace (str_in, '*', '%', [ rfReplaceAll, rfIgnoreCase ]);
rez := stringreplace (rez, '?', '_', [ rfReplaceAll, rfIgnoreCase ]);
result := rez;
end;

function select_printer(Sender: TObject; notfirst: Boolean): Integer;
var i,prc: Integer; pr,pr1,rep: String;
begin
  prc := TBasereport(Sender).Printers.Count;
  result := -1;

  if not mainform.printlist then
  begin
    for i := 0 to prc - 1 do
      begin
        pr1 := TBasereport(Sender).Printers[i];
        mainform.ListBox1.Items.Add(pr1);
      end;
      mainform.printlist := True;
  end;

if notfirst then begin

  mainform.reports.open;
  mainform.reports.refresh;
  rep := mainform.RvProject1.ReportName;
  if mainform.Reports.Locate('name',VarArrayOf([rep]),[])
  then begin
    runsql ('update reports set counter=counter+1 where name="' + rep + '"',mainform.MyConnection1); 
    if mainform.Reportsprinter.AsString = 'Sheet' then pr := config.print_sheet;
    if mainform.Reportsprinter.AsString = 'Zakaz' then pr := config.print_zakaz;
    if mainform.Reportsprinter.AsString = 'Pdf' then pr := config.print_email;
    if mainform.Reportsprinter.AsString = 'Deflt' then pr := config.print_default;
    if mainform.Reportsprinter.AsString = 'A3' then pr := config.print_A3;
    for i := 0 to prc-1 do
    begin
      pr1 := TBasereport(Sender).Printers[i];
      if  pr1 = pr then result := i;
    end;
  end;
  mainform.reports.Close;

end;
end;

procedure savewindow (window: TForm);
var ini: TIniFile;
begin
  ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'view.ini');
  ini.WriteString(window.Name,'top', IntToStr(window.Top) );
  ini.WriteString(window.Name,'left', IntToStr(window.left) );
  ini.WriteString(window.Name,'height', IntToStr(window.height) );
  ini.WriteString(window.Name,'width', IntToStr(window.width) );
  ini.Free;
end;

procedure restorewindow (window: TForm);
var ini: TIniFile;
begin
  ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'view.ini');
  window.Top := StrToInt(ini.ReadString(window.Name,'top', '25' ));
  window.left := StrToInt(ini.ReadString(window.Name,'left', '25' ));
  window.Height := StrToInt(ini.ReadString(window.Name,'height', '250' ));
  window.Width := StrToInt(ini.ReadString(window.Name,'width', '250' ));
  ini.Free;
end;

//*************************************************************************************
//  Процедура заполнения списка поиска в поле TComboBox из данных БД
//

procedure FindComboItems(cb: TCOmboBox; TableName, NameColumn,AddSQL: String; base: TMyConnection; var cbLength: Integer);
var table: TMyQuery; sql: String; i: Integer;
begin
if cbLength = Length(cb.Text) then exit; // нет изменений
if Length (cb.Text) - cbLength > 1
 then begin
   cbLength := Length (cb.Text); exit;
 end;

cbLength := Length (cb.Text);
if cbLength < 1 then exit;

table := TMyQuery.Create(mainform);
table.Connection := base;

sql := 'select ' + NameColumn + ' from ' +  TableName +
       ' where left(' + NameColumn + ',' + IntToStr(length(cb.Text)) + ')="' + cb.Text + '"';
table.SQL.SetText(@sql[1]); table.Open;   table.refresh;

cb.Items.BeginUpdate; cb.Items.Clear;

for i := 1 to table.recordcount do
begin
   cb.Items.Add(table.FieldByName(NameColumn).AsString);
   table.Next;
end;
cb.Items.EndUpdate;
cb.SelStart := length (cb.Text);
cb.SelLength := 0;
table.Close; freeandnil (table);
end;

// ************************************************************************************
// функция возвращает слово номер N из строки S
// слова разделяются разделителем D

function word_from (S: String; D: String; N: integer): String;
var i,p: Integer; bufer: String; flag: boolean;
begin
  result := ''; flag := True;
      bufer := TrimLeft(S);
  for i := 1 to N do
  begin
    if flag then
    begin
      p := AnsiPos (D,bufer);
      if p> 0
        then begin
         result := copy (bufer,1,p-1);
         bufer := copy (bufer,p+1, length(bufer)-p);
        end
        else begin
         result := bufer; flag := false;
        end;

    end
    else result := '';
  end;

end;


// **************************************************************************************
// функция возвращает дату с полным названием месяца
function fulldate (data: TDateTime): string;
var year,month,day: word;
begin
decodedate (data,year,month,day);
result := IntToStr (day) + ' ';
case month of
  1: result := result + 'января ';
  2: result := result + 'февраля ';
  3: result := result + 'марта ';
  4: result := result + 'апреля ';
  5: result := result + 'мая ';
  6: result := result + 'июня ';
  7: result := result + 'июля ';
  8: result := result + 'августа ';
  9: result := result + 'сентября ';
  10: result := result + 'октября ';
  11: result := result + 'ноября ';
  12: result := result + 'декабря ';
end;

result := result + IntToStr (year) + ' г';
end;
//***************************************************************************************
// функция проверки на дублирование номеров форм строгой отчетности
function check_so ( kod: String; ser: String; nom1: String; nom2: String): Boolean;
var checkso: TMyQuery; sql, str_error: String; i,n1,n2: Integer;
begin
  result := true;
  str_error := 'Дублирование номеров:' + chr(13);

  val(nom1,n1,i);   val(nom2,n2,i);

  runsql ('delete from checkso', mainform.MyConnection1);
  runsql ('insert into checkso (nom1,nom2,id) select nom1,nom2,id from realization.prixso ' +
          'where ser="' + ser + '" and pr=' + kod, mainform.MyConnection1);
  runsql ('insert into checkso (nom1,nom2) select nom1,nom2 from tmpso ' +
          'where ser="' + ser + '" and pr=' + kod, mainform.MyConnection1);

  checkso := TMyQuery.Create(mainform);
  checkso.Connection := mainform.MyConnection1;
  sql := 'select nom1,nom2,id from checkso order by nom1';
  checkso.SQL.SetText(@sql[1]);
  checkso.Open;
  for i := 1 to checkso.RecordCount
   do begin
    if checkso.FieldByName('nom1').AsInteger <= n1
     then begin
      result := false;
      if checkso.FieldByName('nom2').AsInteger >= n2
       then str_error := str_error + format( '%7.7d',[n1]) + ' - ' + format( '%7.7d',[n2])
                          + '    № прих. док. : ' + IntToStr(checkso.FieldByName('id').AsInteger) + chr(13)
       else begin
         if checkso.FieldByName('nom2').AsInteger < n1
          then result := true
          else str_error := str_error + format( '%7.7d',[n1]) + ' - '
                                   + format( '%7.7d',[checkso.FieldByName('nom2').AsInteger])
                                   + '    № прих. док. : ' + IntToStr(checkso.FieldByName('id').AsInteger) + chr(13);
       end;
     end
     else begin
      if not checkso.FieldByName('nom1').AsInteger > n2
       then begin
          result := false;
          if checkso.FieldByName('nom2').AsInteger > n2
           then str_error := str_error + format( '%7.7d',[checkso.FieldByName('nom1').AsInteger])
                              + format( '%7.7d',[n2])
                              + '    № прих. док. : ' + IntToStr(checkso.FieldByName('id').AsInteger) + chr(13)
           else str_error := str_error + format( '%7.7d',[checkso.FieldByName('nom1').AsInteger])
                              + format( '%7.7d',[checkso.FieldByName('nom2').AsInteger])
                              + '    № прих. док. : ' + IntToStr(checkso.FieldByName('id').AsInteger) + chr(13);
       end;
     end;
    checkso.Next;
   end;

  if not result then showmessage (str_error);
  
  checkso.Close; freeandnil (checkso);
end;
//***************************************************************************************
function is_mobile (str: String) : String;
var buf: String;
begin
result := '';
if length (str) > 0
then begin
buf:= str;
if buf[1] = '+' then buf := copy (str, 2, length(str)-1);
if length (buf) = 10 then buf := '8' + buf;
if length (buf) = 11  then
  if is_digit (buf)
   then begin
      if buf[1]='8' then buf[1] :='7';
      if buf[1]='7' then result := buf;
   end;
end;
end;

// *****************************************************************************************
function is_latin (str: String): boolean;
// проверка, набрана ли строка в латинском шрифте
var i,c :integer;
begin
  result := True;
  c := length (str);
  for i:=1 to c do
    if not IsDelimiter ('1234567890abcdefjhigklmnopqrstuvwxyzABCDEFJHIGKLMNOPQRSTUVWXYZ',str,i)
    then begin
       result := False;
       exit;
    end;

end;                       


// **************************************************************************************

function is_digit(str:String): boolean;
{ функция проверяет является ли вся строка цифровой
}
var i,c :integer;
begin
  c := length (str);
  if c=0 then is_digit := false else is_digit := true;
  for i:=1 to c do
    if not  IsDelimiter ('1234567890',str,i)
    then begin
       is_digit := False;
       exit;
    end;

end;
// *********************************************************************************************
  procedure corrsklad (sk: String; kod: Integer; delta: Double; connect: TMYConnection; add: boolean);
  var  so,sql,sk1: String; OldSeparator: Char; common_sklad: TMyQuery;
  begin
  OldSeparator := DecimalSeparator;
  DecimalSeparator := '.';
  sql := 'update ' + sk + ' set kol=kol + '  + FloatToStr (delta) + ' where kod=' + IntToStr(kod);
  if add
   then begin
      sk1 := copy ( sk,1, length(sk)-1);
      common_sklad := TMyQuery.Create(mainform);
      common_sklad.Connection := connect;
     if sk[1] = 's' then begin
      sql := 'select kod,kg,name,ed,so from ' + sk1 + ' where kod=' + IntToStr (kod);
      common_sklad.SQL.SetText(@sql[1]);
      common_sklad.Open;
      if common_sklad.FieldByName('so').AsBoolean  then so:='-1' else so := '0';

      runsql ('insert into ' + sk + ' (kod,kg,name,ed,so,kol) ' +
                      ' values (' + IntToStr (kod) +
                                ',' + IntToStr (common_sklad.FieldByName('kg').AsInteger) +
                                ',"' + common_sklad.FieldByName('name').AsString + '"' +
                                ',"' + common_sklad.FieldByName('ed').AsString + '"' +
                                ',' +   so + ',' + FloatToStr (delta) + ') ' +
                      ' on duplicate key update kol=kol+' + FloatToStr (delta), connect);
      common_sklad.Close;
     end
     else begin
      sql := 'select kod,kg,name,ed from ' + sk1 + ' where kod=' + IntToStr (kod);
      common_sklad.SQL.SetText(@sql[1]);
      common_sklad.Open;

      runsql ('insert into ' + sk + ' (kod,kg,name,ed,kol) ' +
                      ' values (' + IntToStr (kod) +
                                ',' + IntToStr (common_sklad.FieldByName('kg').AsInteger) +
                                ',"' + common_sklad.FieldByName('name').AsString + '"' +
                                ',"' + common_sklad.FieldByName('ed').AsString + '"' +
                                ',' + FloatToStr (delta) + ') ' +
                      ' on duplicate key update kol=kol+' + FloatToStr (delta), connect);
      common_sklad.Close;
     end;
      freeandnil (common_sklad);
   end else runsql (sql, connect);
  DecimalSeparator := OldSeparator;
  end;
// *********************************************************************************************
  procedure runsql (sql: String; connect: TMYConnection);
  var script: TMYScript;
  begin
  if connect.Connected then begin
    script := TMYScript.Create(MainForm);
    script.Connection := connect;
    script.SQL.SetText(@sql[1]);
    script.Execute;
    freeandnil (script);
  end;
  end;

// *********************************************************************************************

procedure str_insert_sql ( obj_out: TMyQuery; strin: string; strout: string; var sql_str: string; iskl: integer);
var
 i: integer; sql_select: string; loopcount: integer;
begin
  sql_str := 'insert into ' + strout + ' (';
  sql_select := ' select ';
  loopcount := obj_out.FieldCount - iskl;
  for i := 0 to loopcount - 1 do
    begin
      sql_str := sql_str + obj_out.FieldDefList.FieldDefs[i].Name + ',';
      sql_select := sql_select + obj_out.FieldDefList.FieldDefs[i].Name + ',';
    end;
  sql_str [ length(sql_str) ] := ')';
  sql_select [ length(sql_select) ] := ' ';
  sql_str := sql_str + sql_select + 'from ' + strin;
end;
{***************************************************************************************}
Function MoneyToStr(DD :String):String;
 Type
 TTroyka=Array[1..3] of Byte;
 TMyString=Array[1..19] of String[12];
Var
S,OutS,S2,S3 :String;
k,L,kk :Integer;
Troyka :TTroyka;
V1 :TMyString;
Mb :Byte;
Const
V11 :TMyString=
('один','два','три','четыре','пять','шесть','семь','восемь','девять','десять','одиннадцать',
'двенадцать','тринадцать','четырнадцать','пятнадцать','шестнадцать','семнадцать','восемнадцать','девятнадцать');
V2 :Array[1..8] of String=
('двадцать','тридцать','сорок','пятьдесят','шестьдесят','семьдесят','восемьдесят','девяносто');
V3 :Array[1..9] of String=
('сто','двести','триста','четыреста','пятьсот','шестьсот','семьсот','восемьсот','девятьсот');
M1 :Array[1..13,1..3] of String=(('тысяча','тысячи','тысяч'),
                                ('миллион','миллиона','миллионов'),
                                ('миллиард','миллиарда','миллиардов'),
                                ('триллион','триллиона','триллионов'),
                                ('квадриллион','квадриллиона','квадриллионов'),
                                ('квинтиллион','квинтиллиона','квинтиллионов'),
                                ('секстиллион','секстиллиона','секстиллионов'),
                                ('сентиллион','сентиллиона','сентиллионов'),
                                ('октиллион','октиллиона','октиллионов'),
                                ('нониллион','нониллиона','нониллионов'),
                                ('дециллион','дециллиона','дециллионов'),
                                ('ундециллион','ундециллиона','ундециллионов'),
                                ('додециллион','додециллиона','додециллионов'));
R1 :Array[1..3] of String=('рубль','рубля','рублей');
R2 :Array[1..3] of String=('копейка','копейки','копеек');
  Function TroykaToStr(L :ShortInt;TR :TTroyka):String;
  Var
  S :String;
  Begin
  S:='';
  if Abs(L)=1 then Begin V1[1]:='одна';V1[2]:='две';end else
              Begin V1[1]:='один';V1[2]:='два';end;
  if Troyka[2]=1 then Begin Troyka[2]:=0;Troyka[3]:=10+Troyka[3];end;
  if Troyka[3]<>0 then S:=V1[Troyka[3]];
  if Troyka[2]<>0 then S:=V2[Troyka[2]-1]+' '+S;
  if Troyka[1]<>0 then S:=V3[Troyka[1]]+' '+S;
  if (L>0) and (S<>'') then Case Troyka[3] of
                 1: S:=S+' '+M1[L,1]+' ';
              2..4: S:=S+' '+M1[L,2]+' ';
               else S:=S+' '+M1[L,3]+' ';
              end;
  TroykaToStr:=S;
  End;
Begin
V1:=V11;L:=0;OutS:='';
//kk:=Pos(',',DD);
kk:=Pos(DecimalSeparator,DD);
if kk=0 then S:=DD else S:=Copy(DD,1,kk-1);if S='0' then S2:='' else S2:=S;
 Repeat
 for k:=3 downto 1 do
   if Length(S)>0 then
                  Begin
                  Troyka[k]:=StrToInt(S[Length(S)]);
                  Delete(S,Length(S),1);
                  end else
                  Troyka[k]:=0;
 OutS:=TroykaToStr(L,Troyka)+OutS;
 if L=0 then Mb:=Troyka[3];
 Inc(L);
 Until Length(S)=0;
 case Mb of
    0:if Length(S2)>0 then OutS:=OutS+' '+R1[3]+' ';
    1:OutS:=OutS+' '+R1[1]+' ';
 2..4:OutS:=OutS+' '+R1[2]+' ';
 else OutS:=OutS+' '+R1[3]+' ';
 end;S2:='';
 if kk<>0 then
 Begin
   DD:=Copy(DD,kk+1,2);if Length(DD)=1 then DD:=DD+'0';
   k:=StrToInt(DD);
   Troyka[1]:=0;Troyka[2]:=k div 10;Troyka[3]:=k mod 10;S2:=TroykaToStr(-1,Troyka);
   case Troyka[3] of
      0:if Troyka[2]=0 then S:='' else S:=R2[3];
      1:S:=R2[1];
   2..4:S:=R2[2];
   else S:=R2[3];
   end;
 end;
 if OutS <> '' then begin
 S3 := OutS[1];
 S3 := AnsiUpperCase (S3);
 OutS[1] := S3 [1];
 end
 else OutS := 'ноль рублей ';

 if k <> 0  then begin
 if k > 9
  then MoneyToStr:=OutS+IntToStr(k)+' '+S // если копейки нужны цифрой-эту строку раскоментировать
  else MoneyToStr:=OutS+'0'+IntToStr(k)+' '+S
  end
 else MoneyToStr:=OutS + ' 00 копеек';
// MoneyToStr:=OutS+S2+' '+S; // а эту закоментировать

End;
//********************************************************************************************
// сворачивание в имени двойных и одинарных кавычек в угловые кавычки
// (для корректной работы с Мускулом
//********************************************************************************************
function uni_name (str_in: string): string;
var Flags: TReplaceFlags; L,I: Integer; zakr: boolean; buf: string;
begin
  buf := str_in;
  zakr := False;
  L := length (buf);
  Flags := [rfReplaceAll, rfIgnoreCase];
  buf := StringReplace (str_in, '''', '"', Flags);
  buf := StringReplace (str_in, '"', '""', Flags); // так двойные кавычки оставляем в строке
 { for I := 1 to L do
  if buf[I] = '"'
   then begin
      if zakr  then buf[I] := chr (187)         // а это замена двойных кавычек на угловые
        else buf[I] := chr (171);
      zakr := not zakr;
   end;
  }
   uni_name := trim(buf);
end;

function TOpenOfficeConnect: boolean;
begin
   if VarIsEmpty(OO) then
      OO := CreateOleObject('com.sun.star.ServiceManager');
   Result := not (VarIsEmpty(OO) or VarIsNull(OO));
end;

function TOpenOfficeOpenDocument(const FileName:string):
                                                     boolean;
var
 Desktop: Variant;
 VariantArray: Variant;
begin
 Desktop := OO.CreateInstance('com.sun.star.frame.Desktop');
 VariantArray := VarArrayCreate([0, 0], varVariant);
 VariantArray[0] := TOpenOfficeMakePropertyValue('FilterName',
                                                'MS Excel 97');
 Document := Desktop.LoadComponentFromURL('file://localhost/'+FileName , '_blank',
                                                0,VariantArray);
 Result := not (VarIsEmpty(Document) or VarIsNull(Document));
end;

function TOpenOfficeCreateDocument: boolean;
var
   Desktop: Variant;
begin
   Desktop := OO.createInstance('com.sun.star.frame.Desktop');
   Document := Desktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0,
                  VarArrayCreate([0, -1], varVariant));
   Result := not (VarIsEmpty(Document) or VarIsNull(Document));
end;

function TOpenOfficeMakePropertyValue(PropertyName,
                                PropertyValue:string):variant;
var
 Structure: variant;
begin
 Structure :=
       OO.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
 Structure.Name := PropertyName;
 Structure.Value := PropertyValue;
 Result := Structure;
end;
//***************************************************************************************
function DateExists(Date: string; Separator: char): Boolean;
var
  OldDateSeparator: Char;
begin
  Result := True;
  OldDateSeparator := DateSeparator;
  DateSeparator := Separator;
  try
    try
      StrToDate(Date);
    except
      Result := False;
    end;
  finally
    DateSeparator := OldDateSeparator;
  end;
end;

function GetInetFile(fileURL: string; FileName: String): boolean;
const BufferSize = 1024;
var hSession, hURL: HInternet;
Buffer: array[1..BufferSize] of Byte;
BufferLen: DWORD;
f: File;
sAppName: string;
begin
  GetInetFile := False;
   Result:=False;
   sAppName := ExtractFileName(Application.ExeName);
   hSession := InternetOpen(PChar(sAppName), INTERNET_OPEN_TYPE_DIRECT,
         nil, nil, 0);
   try
      hURL := InternetOpenURL(hSession,
      PChar(fileURL),nil,0,0,0);
      if hURL <> nil then
      try
         AssignFile(f, FileName);
         Rewrite(f,1);
         repeat
            InternetReadFile(hURL, @Buffer, SizeOf(Buffer), BufferLen);
            BlockWrite(f, Buffer, BufferLen)
         until BufferLen = 0;
         CloseFile(f);
         Result:=True;
         GetInetFile := True;
      finally
      InternetCloseHandle(hURL);
      end
   finally
   InternetCloseHandle(hSession);
   end
end;

 function IsFormOpen(const FormName : string): Boolean;
 var
   i: Integer;
 begin
   Result := False;
   for i := Screen.FormCount - 1 DownTo 0 do
     if (Screen.Forms[i].Name = FormName) then
     begin
       Result := True;
       Break;
     end;
 end;

end.
