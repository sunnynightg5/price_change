unit SabzPricing;

{ ---------------------------------------------------------------------------
  لایه داده و منطق قیمت‌گذاری استودیو سبز
  دسترسی به جدول Price و person در پایگاه داده SABZ
  --------------------------------------------------------------------------- }

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  System.StrUtils, System.Math, Data.DB, Data.Win.ADODB;

type
  TPriceCategory = (pcArtReprint, pcArtNew, pcFrame, pcFace,
    pcItalianAlbum, pcCeremony, pcPrintCost, pcFrameCost, pcLaminate,
    pcFile, pcOther);

const
  PriceCategoryNames: array[TPriceCategory] of string = (
    'عکس هنری - چاپ مجدد',
    'عکس هنری - عکس جدید (فقط چاپ)',
    'شاسی و لمینت تنها',
    'هزینه چهره اضافه (روتوش)',
    'آلبوم ایتالیایی عروس و داماد',
    'عکس مراسم عروس و داماد',
    'قیمت تمام‌شده چاپ (خرید ما)',
    'قیمت تمام‌شده شاسی (خرید ما)',
    'قیمت تمام‌شده لمینت (خرید ما)',
    'فقط فایل',
    'سایر'
    );

  SabzConnStr =
    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;' +
    'Initial Catalog=SABZ;Data Source=.';

type
  TPriceItem = class
  private
    FCode: string;
    FPrice: Int64;
    FCount: Integer;
    FTakhfif: string;
    FCategory: TPriceCategory;
    FModified: Boolean;
  public
    property Code: string read FCode write FCode;
    property Price: Int64 read FPrice write FPrice;
    property Count: Integer read FCount write FCount;
    property Takhfif: string read FTakhfif write FTakhfif;
    property Category: TPriceCategory read FCategory write FCategory;
    property Modified: Boolean read FModified write FModified;
    function DisplayName: string;
  end;

  TPerson = class
  public
    Id: Int64;
    Name: string;
    Sex: string;
    Mobile: string;
    NationalCode: string;
    Job: string;
    OrderCount: Integer;
    function SexLabel: string;
  end;

  TSabzData = class
  private
    FConn: TADOConnection;
    FQ: TADOQuery;
    FLastError: string;
    FConnStr: string;
  public
    constructor Create;
    destructor Destroy; override;
    function Connect: Boolean;
    procedure Disconnect;
    function IsConnected: Boolean;
    function LoadPrices(AList: TObjectList<TPriceItem>): Boolean;
    function SavePrice(const Code: string; NewPrice: Int64): Boolean;
    function InsertPrice(const Code: string; NewPrice: Int64): Boolean;
    function SearchPeople(const Term: string; Limit: Integer;
      AList: TObjectList<TPerson>): Boolean;
    function LoadAllPeople(AList: TObjectList<TPerson>): Boolean;
    function CountPeople: Integer;
    function CountPriceRows: Integer;
    property LastError: string read FLastError;
    property Conn: TADOConnection read FConn;
    { رشته اتصال؛ از فایل تنظیمات قابل تغییر است }
    property ConnStr: string read FConnStr write FConnStr;
  end;

function ClassifyCode(const Code: string): TPriceCategory;
function SizePartOf(const Code: string): string;
function PriceDisplayName(const Code: string): string;
function FormatMoney(v: Int64): string;
function FormatToman(v: Int64): string;
function ParseMoney(const S: string): Int64;
function Levenshtein(const S1, S2: string): Integer;
function ToLatinDigits(const S: string): string;
function ToPersianDigits(const S: string): string;
function GregorianToJalali(gy, gm, gd: Integer): string;
function JalaliToday: string;
function FindPrice(AList: TObjectList<TPriceItem>; const Code: string): TPriceItem;
function NumberToPersianWords(v: Int64): string;
function NormalizeText(const S: string; KeepSpaces: Boolean = False): string;
function FuzzyDistance(const Term, Text: string): Integer;
function IsDigitsOnly(const S: string): Boolean;

implementation

{ ------------------------------- helpers --------------------------------- }

function ClassifyCode(const Code: string): TPriceCategory;
var
  c: string;
begin
  c := Trim(Code);
  if c = '' then
    Exit(pcOther);
  if StartsText('f_', c) then
    Exit(pcFile);
  if StartsText('p', c) then
    Exit(pcPrintCost);
  if StartsText('s', c) then
    Exit(pcFrameCost);
  if StartsText('l', c) and (Length(c) > 1) and CharInSet(c[2], ['0' .. '9']) then
    Exit(pcLaminate);
  if EndsText('i', c) then
    Exit(pcItalianAlbum);
  if EndsText('a', c) then
    Exit(pcCeremony);
  if EndsText('f', c) then
    Exit(pcFace);
  if EndsText('_', c) then
    Exit(pcFrame);
  if EndsText('n', c) then
    Exit(pcArtNew);
  if Pos('*', c) > 0 then
    Exit(pcArtReprint);
  Result := pcOther;
end;

function ExtractSize(const Code: string): string;
var
  c: string;
begin
  c := Trim(Code);
  if StartsText('f_', c) then
    Exit('');
  if (Length(c) > 0) and (CharInSet(c[1], ['p', 'P', 's', 'S', 'l', 'L'])) and
    (Pos('*', c) > 0) then
    c := Copy(c, 2, MaxInt);
  if (Length(c) > 0) and
    (CharInSet(c[Length(c)], ['n', 'N', 'f', 'F', 'i', 'I', 'a', 'A', '_', 'p', 'P']))
  then
    c := Copy(c, 1, Length(c) - 1);
  Result := Trim(c);
end;

function SizePartOf(const Code: string): string;
begin
  Result := ExtractSize(Code);
end;

function PriceDisplayName(const Code: string): string;
var
  c, sz: string;
begin
  c := Trim(Code);
  if SameText(c, 'cd') then
    Exit('سی‌دی (CD)');
  if SameText(c, 'cd_just') then
    Exit('سی‌دی فقط');
  if SameText(c, 'new') then
    Exit('عکس جدید (سفارش ورودی)');
  if SameText(c, 'nose') then
    Exit('نوز / اصلاح بینی');
  if SameText(c, 'Lottery') then
    Exit('قرعه‌کشی');
  if SameText(c, '24') then
    Exit('بسته ۲۴ عددی');
  if SameText(c, 'f_rotosh') then
    Exit('فایل با روتوش');
  if SameText(c, 'f_norotosh') then
    Exit('فایل بدون روتوش');
  if SameText(c, 'f_print') then
    Exit('فایل بعد از چاپ');
  if SameText(c, 'f_jest') then
    Exit('فایل ژست چاپ‌نشده');

  sz := ExtractSize(c);
  if sz = '' then
    sz := c;
  sz := StringReplace(sz, '*', ' × ', [rfReplaceAll]);
  sz := #$202A + sz + #$202C;

  case ClassifyCode(c) of
    pcArtReprint:
      Result := 'عکس هنری - چاپ مجدد ' + sz;
    pcArtNew:
      Result := 'عکس هنری - عکس جدید (فقط چاپ) ' + sz;
    pcFrame:
      Result := 'شاسی و لمینت تنها ' + sz;
    pcFace:
      Result := 'هزینه هر چهره اضافه (روتوش) ' + sz;
    pcItalianAlbum:
      Result := 'آلبوم ایتالیایی عروس و داماد ' + sz;
    pcCeremony:
      Result := 'عکس مراسم عروس و داماد ' + sz;
    pcPrintCost:
      Result := 'هزینه چاپ - تمام‌شده برای ما ' + sz;
    pcFrameCost:
      Result := 'هزینه شاسی - تمام‌شده برای ما ' + sz;
    pcLaminate:
      Result := 'هزینه لمینت - تمام‌شده برای ما ' + sz;
    pcFile:
      Result := 'فایل ' + c;
  else
    Result := 'کد: ' + c;
  end;
end;

function FormatMoney(v: Int64): string;
begin
  Result := FormatFloat('#,##0', v);
end;

function FormatToman(v: Int64): string;
begin
  Result := FormatFloat('#,##0', v) + ' تومان';
end;

function ToLatinDigits(const S: string): string;
var
  i: Integer;
  ch: Char;
begin
  { ارقام فارسی (۰-۹) و عربی (٠-٩) هر دو به لاتین تبدیل می‌شوند }
  Result := '';
  for i := 1 to Length(S) do
  begin
    ch := S[i];
    if (ch >= #$06F0) and (ch <= #$06F9) then
      Result := Result + Chr(Ord('0') + Ord(ch) - $06F0)
    else if (ch >= #$0660) and (ch <= #$0669) then
      Result := Result + Chr(Ord('0') + Ord(ch) - $0660)
    else
      Result := Result + ch;
  end;
end;

function IsDigitsOnly(const S: string): Boolean;
var
  ch: Char;
  t: string;
begin
  t := ToLatinDigits(Trim(S));
  Result := t <> '';
  for ch in t do
    if not CharInSet(ch, ['0' .. '9', ' ', '-', '+']) then
      Exit(False);
end;

function NormalizeText(const S: string; KeepSpaces: Boolean): string;
var
  ch: Char;
begin
  { یکسان‌سازی برای جستجو: ی/ک عربی، ارقام، حروف کوچک، × و x به * }
  Result := '';
  for ch in LowerCase(ToLatinDigits(Trim(S))) do
    case ch of
      #$064A, #$0649:
        Result := Result + #$06CC;
      #$0643:
        Result := Result + #$06A9;
      #$0629:
        Result := Result + #$0647;
      #$00D7, 'x':
        Result := Result + '*';
      #$200C, #$200F, #$0640:
        ;
      ' ':
        if KeepSpaces then
          Result := Result + ch;
    else
      Result := Result + ch;
    end;
end;

function FuzzyDistance(const Term, Text: string): Integer;
var
  tw, xw: TArray<string>;
  t, x: string;
  best, d: Integer;
begin
  { برای هر کلمه جستجو نزدیک‌ترین کلمه متن پیدا و فاصله‌ها جمع می‌شود.
    پیشوند بودن کلمه، تطابق کامل حساب می‌شود. }
  tw := NormalizeText(Term, True).Split([' '], TStringSplitOptions.ExcludeEmpty);
  xw := NormalizeText(Text, True).Split([' '], TStringSplitOptions.ExcludeEmpty);
  if (Length(tw) = 0) or (Length(xw) = 0) then
    Exit(MaxInt div 2);
  Result := 0;
  for t in tw do
  begin
    best := MaxInt div 2;
    for x in xw do
    begin
      if StartsStr(t, x) then
        d := 0
      else
        d := Levenshtein(t, x);
      if d < best then
        best := d;
    end;
    Inc(Result, best);
  end;
end;

function NumberToPersianWords(v: Int64): string;
const
  Ones: array [0 .. 19] of string = ('', 'یک', 'دو', 'سه', 'چهار', 'پنج',
    'شش', 'هفت', 'هشت', 'نه', 'ده', 'یازده', 'دوازده', 'سیزده', 'چهارده',
    'پانزده', 'شانزده', 'هفده', 'هجده', 'نوزده');
  Tens: array [2 .. 9] of string = ('بیست', 'سی', 'چهل', 'پنجاه', 'شصت',
    'هفتاد', 'هشتاد', 'نود');
  Hundreds: array [1 .. 9] of string = ('صد', 'دویست', 'سیصد', 'چهارصد',
    'پانصد', 'ششصد', 'هفتصد', 'هشتصد', 'نهصد');
  Scales: array [0 .. 6] of string = ('', 'هزار', 'میلیون', 'میلیارد',
    'هزار میلیارد', 'میلیون میلیارد', 'میلیارد میلیارد');

  function Below1000(n: Integer): string;
  begin
    Result := '';
    if n >= 100 then
    begin
      Result := Hundreds[n div 100];
      n := n mod 100;
    end;
    if n = 0 then
      Exit;
    if Result <> '' then
      Result := Result + ' و ';
    if n < 20 then
      Result := Result + Ones[n]
    else
    begin
      Result := Result + Tens[n div 10];
      if n mod 10 <> 0 then
        Result := Result + ' و ' + Ones[n mod 10];
    end;
  end;

var
  neg: Boolean;
  scale, g: Integer;
  part: string;
begin
  if v = 0 then
    Exit('صفر');
  neg := v < 0;
  if neg then
    v := -v;
  Result := '';
  scale := 0;
  while (v > 0) and (scale <= High(Scales)) do
  begin
    g := v mod 1000;
    v := v div 1000;
    if g > 0 then
    begin
      part := Below1000(g);
      if Scales[scale] <> '' then
        part := part + ' ' + Scales[scale];
      if Result = '' then
        Result := part
      else
        Result := part + ' و ' + Result;
    end;
    Inc(scale);
  end;
  if neg then
    Result := 'منفی ' + Result;
end;

function ParseMoney(const S: string): Int64;
var
  t: string;
  i: Integer;
  digits: string;
begin
  t := ToLatinDigits(S);
  digits := '';
  for i := 1 to Length(t) do
    if CharInSet(t[i], ['0' .. '9']) then
      digits := digits + t[i];
  if digits = '' then
    Result := 0
  else
    Result := StrToInt64Def(digits, 0);
end;

function ToPersianDigits(const S: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(S) do
    case S[i] of
      '0': Result := Result + '۰';
      '1': Result := Result + '۱';
      '2': Result := Result + '۲';
      '3': Result := Result + '۳';
      '4': Result := Result + '۴';
      '5': Result := Result + '۵';
      '6': Result := Result + '۶';
      '7': Result := Result + '۷';
      '8': Result := Result + '۸';
      '9': Result := Result + '۹';
    else
      Result := Result + S[i];
    end;
end;

function Levenshtein(const S1, S2: string): Integer;
var
  d: array of Integer;
  i, j, cost, n, m, lastDiag, oldDiag: Integer;
begin
  n := Length(S1);
  m := Length(S2);
  if n = 0 then
    Exit(m);
  if m = 0 then
    Exit(n);
  SetLength(d, m + 1);
  for j := 0 to m do
    d[j] := j;
  for i := 1 to n do
  begin
    lastDiag := d[0];
    Inc(d[0]);
    for j := 1 to m do
    begin
      oldDiag := d[j];
      if S1[i] = S2[j] then
        cost := 0
      else
        cost := 1;
      d[j] := Min(Min(d[j] + 1, d[j - 1] + 1), lastDiag + cost);
      lastDiag := oldDiag;
    end;
  end;
  Result := d[m];
end;

function GregorianToJalali(gy, gm, gd: Integer): string;
const
  g_d_m: array [1 .. 12] of Integer = (0, 31, 59, 90, 120, 151, 181, 212,
    243, 273, 304, 334);
var
  gy2, days, jy, jm, jd: Integer;
begin
  gy2 := gy;
  if gm > 2 then
    Inc(gy2);
  days := 355666 + (365 * gy) + ((gy2 + 3) div 4) - ((gy2 + 99) div 100) +
    ((gy2 + 399) div 400) + gd + g_d_m[gm];
  jy := -1595 + (33 * (days div 12053));
  days := days mod 12053;
  jy := jy + (4 * (days div 1461));
  days := days mod 1461;
  if days > 365 then
  begin
    jy := jy + ((days - 1) div 365);
    days := (days - 1) mod 365;
  end;
  if days < 186 then
  begin
    jm := 1 + (days div 31);
    jd := 1 + (days mod 31);
  end
  else
  begin
    jm := 7 + ((days - 186) div 30);
    jd := 1 + ((days - 186) mod 30);
  end;
  Result := Format('%.4d/%.2d/%.2d', [jy, jm, jd]);
end;

function JalaliToday: string;
var
  y, m, d: Word;
begin
  DecodeDate(Now, y, m, d);
  Result := GregorianToJalali(y, m, d);
end;

function FindPrice(AList: TObjectList<TPriceItem>; const Code: string): TPriceItem;
var
  it: TPriceItem;
begin
  Result := nil;
  if AList = nil then
    Exit;
  for it in AList do
    if SameText(Trim(it.Code), Trim(Code)) then
      Exit(it);
end;

{ ------------------------------ TPriceItem ------------------------------- }

function TPriceItem.DisplayName: string;
begin
  Result := PriceDisplayName(FCode);
end;

{ -------------------------------- TPerson -------------------------------- }

function TPerson.SexLabel: string;
begin
  if SameText(Trim(Sex), 'man') or SameText(Trim(Sex), 'مرد') or
    (Trim(Sex) = 'آقا') then
    Result := 'آقا'
  else if SameText(Trim(Sex), 'woman') or SameText(Trim(Sex), 'زن') or
    (Trim(Sex) = 'خانم') then
    Result := 'خانم'
  else
    Result := Trim(Sex);
end;

{ ------------------------------- TSabzData ------------------------------- }

constructor TSabzData.Create;
begin
  inherited Create;
  FConn := TADOConnection.Create(nil);
  FConn.LoginPrompt := False;
  FConnStr := SabzConnStr;
  FConn.ConnectionString := FConnStr;
  FConn.ConnectionTimeout := 8;
  FQ := TADOQuery.Create(nil);
  FQ.Connection := FConn;
end;

destructor TSabzData.Destroy;
begin
  try
    if FConn.Connected then
      FConn.Connected := False;
  except
  end;
  FQ.Free;
  FConn.Free;
  inherited Destroy;
end;

function TSabzData.Connect: Boolean;
begin
  FLastError := '';
  try
    if FConn.Connected then
      FConn.Connected := False;
    FConn.ConnectionString := FConnStr;
    FConn.Connected := True;
    Result := True;
  except
    on E: Exception do
    begin
      FLastError := E.Message;
      Result := False;
    end;
  end;
end;

procedure TSabzData.Disconnect;
begin
  try
    FConn.Connected := False;
  except
  end;
end;

function TSabzData.IsConnected: Boolean;
begin
  Result := FConn.Connected;
end;

function TSabzData.LoadPrices(AList: TObjectList<TPriceItem>): Boolean;
var
  it: TPriceItem;
begin
  Result := False;
  FLastError := '';
  try
    FQ.Close;
    FQ.SQL.Text := 'SELECT size, [count], price, takhfif FROM Price';
    FQ.Open;
    AList.Clear;
    while not FQ.Eof do
    begin
      it := TPriceItem.Create;
      it.Code := Trim(FQ.Fields[0].AsString);
      it.Count := FQ.Fields[1].AsInteger;
      it.Price := ParseMoney(FQ.Fields[2].AsString);
      it.Takhfif := Trim(FQ.Fields[3].AsString);
      it.Category := ClassifyCode(it.Code);
      it.Modified := False;
      AList.Add(it);
      FQ.Next;
    end;
    Result := True;
  except
    on E: Exception do
      FLastError := E.Message;
  end;
end;

function TSabzData.SavePrice(const Code: string; NewPrice: Int64): Boolean;
begin
  Result := False;
  FLastError := '';
  try
    FQ.Close;
    FQ.SQL.Text := 'UPDATE Price SET price = :pr WHERE RTRIM(size) = :sz';
    FQ.Parameters.ParamByName('pr').Value := IntToStr(NewPrice);
    FQ.Parameters.ParamByName('sz').Value := Trim(Code);
    FQ.ExecSQL;
    Result := True;
  except
    on E: Exception do
      FLastError := E.Message;
  end;
end;

function TSabzData.InsertPrice(const Code: string; NewPrice: Int64): Boolean;
begin
  Result := False;
  FLastError := '';
  try
    FQ.Close;
    FQ.SQL.Text :=
      'INSERT INTO Price(size, [count], price, takhfif) VALUES(:sz, 1, :pr, ''0'')';
    FQ.Parameters.ParamByName('sz').Value := Trim(Code);
    FQ.Parameters.ParamByName('pr').Value := IntToStr(NewPrice);
    FQ.ExecSQL;
    Result := True;
  except
    on E: Exception do
      FLastError := E.Message;
  end;
end;

function TSabzData.SearchPeople(const Term: string; Limit: Integer;
  AList: TObjectList<TPerson>): Boolean;
var
  p: TPerson;
  pat: string;
begin
  Result := False;
  FLastError := '';
  try
    FQ.Close;
    FQ.SQL.Text :=
      'SELECT TOP ' + IntToStr(Limit) +
      ' Id_p, sex, Name, National_Code, Mobile, job, [count] FROM person ' +
      'WHERE (Name LIKE :t1) OR (Name LIKE :t4) ' +
      'OR (CAST(Mobile AS nvarchar(20)) LIKE :t2) ' +
      'OR (CAST(National_Code AS nvarchar(20)) LIKE :t3) ORDER BY Id_p DESC';
    pat := '%' + Term + '%';
    FQ.Parameters.ParamByName('t1').Value := pat;
    FQ.Parameters.ParamByName('t2').Value := pat;
    FQ.Parameters.ParamByName('t3').Value := pat;
    { نام‌هایی که با «ي/ك» عربی ثبت شده‌اند هم پیدا شوند }
    pat := StringReplace(pat, #$06CC, #$064A, [rfReplaceAll]);
    pat := StringReplace(pat, #$06A9, #$0643, [rfReplaceAll]);
    FQ.Parameters.ParamByName('t4').Value := pat;
    FQ.Open;
    AList.Clear;
    while not FQ.Eof do
    begin
      p := TPerson.Create;
      p.Id := FQ.Fields[0].AsLargeInt;
      p.Sex := Trim(FQ.Fields[1].AsString);
      p.Name := Trim(FQ.Fields[2].AsString);
      p.NationalCode := FQ.Fields[3].AsString;
      p.Mobile := FQ.Fields[4].AsString;
      p.Job := Trim(FQ.Fields[5].AsString);
      p.OrderCount := FQ.Fields[6].AsInteger;
      AList.Add(p);
      FQ.Next;
    end;
    Result := True;
  except
    on E: Exception do
      FLastError := E.Message;
  end;
end;

function TSabzData.LoadAllPeople(AList: TObjectList<TPerson>): Boolean;
var
  p: TPerson;
begin
  Result := False;
  FLastError := '';
  try
    FQ.Close;
    FQ.SQL.Text :=
      'SELECT Id_p, sex, Name, National_Code, Mobile, job, [count] FROM person';
    FQ.Open;
    AList.Clear;
    while not FQ.Eof do
    begin
      p := TPerson.Create;
      p.Id := FQ.Fields[0].AsLargeInt;
      p.Sex := Trim(FQ.Fields[1].AsString);
      p.Name := Trim(FQ.Fields[2].AsString);
      p.NationalCode := FQ.Fields[3].AsString;
      p.Mobile := FQ.Fields[4].AsString;
      p.Job := Trim(FQ.Fields[5].AsString);
      p.OrderCount := FQ.Fields[6].AsInteger;
      AList.Add(p);
      FQ.Next;
    end;
    Result := True;
  except
    on E: Exception do
      FLastError := E.Message;
  end;
end;

function TSabzData.CountPeople: Integer;
begin
  Result := -1;
  try
    FQ.Close;
    FQ.SQL.Text := 'SELECT COUNT(*) FROM person';
    FQ.Open;
    if not FQ.Eof then
      Result := FQ.Fields[0].AsInteger;
  except
    Result := -1;
  end;
end;

function TSabzData.CountPriceRows: Integer;
begin
  Result := -1;
  try
    FQ.Close;
    FQ.SQL.Text := 'SELECT COUNT(*) FROM Price';
    FQ.Open;
    if not FQ.Eof then
      Result := FQ.Fields[0].AsInteger;
  except
    Result := -1;
  end;
end;

end.
