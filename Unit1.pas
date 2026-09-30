unit Unit1;

{ ---------------------------------------------------------------------------
  سامانه مدیریت قیمت استودیو عکاسی سبز
  - داشبورد
  - جدول قیمت (خوانا)
  - ویرایش قیمت‌ها (اتصال به جدول Price)
  - محاسبه سفارش (سبد سفارش، تخفیف درصدی، فاکتور با مبلغ به حروف)
  - هزینه‌های تمام‌شده
  - مشتریان (جستجوی فازی، مقاوم به غلط تایپی و ی/ک عربی)
  - تاریخچه تغییر قیمت‌ها (PriceLog.csv کنار برنامه)
  - میان‌برها: F1..F6 صفحات، Ctrl+S ذخیره، Ctrl+F جستجو، Ctrl+R بارگذاری
  --------------------------------------------------------------------------- }

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI,
  System.SysUtils, System.Variants, System.Classes, System.Types,
  System.Win.ComObj, System.IniFiles, System.Generics.Defaults,
  System.Generics.Collections, System.StrUtils, System.Math, System.UITypes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.Grids, Vcl.Clipbrd,
  Data.DB, Data.Win.ADODB,
  SabzPricing;

type
  TCartLine = class
    Code: string;
    Title: string;
    Qty: Integer;
    UnitPrice: Int64;
    Cost: Int64;
    function Total: Int64;
  end;

  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    FData: TSabzData;
    FPrices: TObjectList<TPriceItem>;
    FIndex: TDictionary<string, TPriceItem>;
    FCart: TObjectList<TCartLine>;
    FPages: array [0 .. 5] of TPanel;
    FNav: array [0 .. 5] of TPanel;
    FExcelPath: string;
    FShopName: string;
    FShopPhone: string;
    FShopAddress: string;
    FChkShop: TCheckBox;
    FEdInvDate: TEdit;
    FCustomerMobile: string;
    FGridLog: TStringGrid;
    FGridPeople: TStringGrid;
    FEdPeopleSearch: TEdit;
    FLblPeopleInfo: TLabel;
    FPeopleTimer: TTimer;
    FPeopleDb: TObjectList<TPerson>;
    FAllPeople: TObjectList<TPerson>;
    FAllPeopleLoaded: Boolean;
    FPeopleView: TList<TPerson>;
    FPeopleDist: TList<Integer>;
    FCurrentPage: Integer;
    FUpdating: Boolean;
    FLblConn: TLabel;
    FBtnReconnect: TPanel;
    FGridMatrix: TStringGrid;
    FGridPrice: TStringGrid;
    FEdPriceSearch: TEdit;
    FCbCat: TComboBox;
    FBtnSave: TPanel;
    FLblChanges: TLabel;
    FEdPercent: TEdit;
    FCbRound: TComboBox;
    FCbRoundMode: TComboBox;
    FBtnApplyPct: TPanel;
    FBtnUndoPct: TPanel;
    FLblPctInfo: TLabel;
    FOriginal: TDictionary<string, Int64>;
    FPctBackup: TDictionary<string, Int64>;
    FCostDraft: TDictionary<string, Int64>;
    FGridCost: TStringGrid;
    FBtnCostSave: TPanel;
    FLblCostInfo: TLabel;
    FGridCatalog: TStringGrid;
    FEdCatSearch: TEdit;
    FGridCart: TStringGrid;
    FEdDisc: TEdit;
    FLblSubtotal: TLabel;
    FLblCost: TLabel;
    FLblProfit: TLabel;
    FLblPayable: TLabel;
    FEdCustomer: TEdit;
    FLblStat: array [0 .. 3] of TLabel;
    FStatCards: array [0 .. 3] of TPanel;
    FCalcOuter: TPanel;
    FCalcCatCard: TPanel;
    FStatus: TLabel;
    procedure BuildHeader;
    procedure BuildNav;
    procedure BuildFooter;
    procedure BuildDashboard;
    procedure BuildMatrixPage;
    procedure BuildPricePage;
    procedure BuildCalcPage;
    procedure BuildCostPage;
    procedure FillCostGrid;
    procedure GridCostSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GridCostSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
    procedure BtnCostSaveClick(Sender: TObject);
    procedure BtnGoCostClick(Sender: TObject);
    procedure BtnFromExcelClick(Sender: TObject);
    procedure BtnToExcelClick(Sender: TObject);
    procedure ShowPage(Index: Integer);
    procedure ReloadPrices;
    procedure BuildOriginalMap;
    procedure FillMatrix;
    procedure FillPriceGrid;
    procedure FillCatalog;
    function PriceOf(const Code: string): Int64;
    procedure RefreshCart(ASelect: Integer = -1);
    procedure RecalcCart;
    procedure UpdateConnStatus;
    procedure UpdateStatCards;
    procedure LayoutStats(Sender: TObject);
    procedure LayoutCalc(Sender: TObject);
    procedure SetStatus(const S: string);
    procedure UpdateChangesLabel;
    procedure BtnEnter(Sender: TObject);
    procedure BtnLeave(Sender: TObject);
    procedure NavClick(Sender: TObject);
    procedure NavEnter(Sender: TObject);
    procedure NavLeave(Sender: TObject);
    procedure BtnReconnectClick(Sender: TObject);
    procedure EdPriceSearchChange(Sender: TObject);
    procedure CbCatChange(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure ApplyPctClick(Sender: TObject);
    procedure UndoPctClick(Sender: TObject);
    function IsPriceModified(it: TPriceItem): Boolean;
    function PriceMatchesFilter(it: TPriceItem): Boolean;
    procedure GridPriceSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GridPriceSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
    procedure GridPriceDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure GridMatrixDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure EdCatSearchChange(Sender: TObject);
    procedure GridCatDblClick(Sender: TObject);
    procedure GridCatDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure BtnAddClick(Sender: TObject);
    procedure BtnRemoveClick(Sender: TObject);
    procedure BtnClearCartClick(Sender: TObject);
    procedure GridCartSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GridCartSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
    procedure GridCartDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure EdDiscChange(Sender: TObject);
    procedure BtnCopyInvoiceClick(Sender: TObject);
    procedure BtnPrintInvoiceClick(Sender: TObject);
    procedure BtnSaveInvoiceClick(Sender: TObject);
    procedure BtnOpenFacturesClick(Sender: TObject);
    procedure BtnSvcBothClick(Sender: TObject);
    procedure BtnSvcFrameClick(Sender: TObject);
    procedure BtnSvcLamClick(Sender: TObject);
    procedure AddService(Kind: Integer);
    function ServiceSize: string;
    function EstimatedCostOf(const Code: string): Int64;
    function InvoiceText(out DocNo: string; out Revenue, Cost, Profit,
      Discount, Payable: Int64): string;
    function InvoiceHTML(const DocNo: string): string;
    procedure BtnGoEditClick(Sender: TObject);
    procedure BtnGoCalcClick(Sender: TObject);
    procedure BtnGoMatrixClick(Sender: TObject);
    { ---- افزوده‌های نسخه ۳ ---- }
    procedure LoadSettings;
    procedure SaveSettings;
    function ResolveExcelPath: Boolean;
    procedure RebuildIndex;
    function FindItem(const Code: string): TPriceItem;
    function PendingPriceCount: Integer;
    function ConfirmPending(IncludeCost: Boolean): Boolean;
    function DoSavePrices: Boolean;
    function DoSaveCosts: Boolean;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure LogPriceChange(const Code: string; OldV, NewV: Int64;
      const Source: string);
    procedure FillLogGrid;
    procedure GridLogDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure GridCostDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    function DiscountValue(Subtotal: Int64): Int64;
    procedure BuildPeoplePage;
    procedure EdPeopleSearchChange(Sender: TObject);
    procedure PeopleTimerTimer(Sender: TObject);
    procedure RunPeopleSearch;
    procedure FillPeopleGrid;
    procedure PickSelectedPerson(Sender: TObject);
    procedure GridPeopleDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure BtnGoPeopleClick(Sender: TObject);
    procedure EdSearchKeyPress(Sender: TObject; var Key: Char);
    procedure EdCustomerChange(Sender: TObject);
    procedure BtnInvTodayClick(Sender: TObject);
    procedure EdInvDateEnter(Sender: TObject);
    function InvoiceDateText: string;
  public
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

var
  C_BG, C_CARD, C_NAVY, C_ACCENT, C_GOLD, C_TEXT, C_MUTED, C_BORDER, C_ALT,
    C_SUCCESS, C_DANGER, C_SEL, C_ACCENT_L: TColor;

const
  NavTitles: array [0 .. 5] of string = ('داشبورد', 'جدول قیمت', 'ویرایش قیمت',
    'محاسبه سفارش', 'هزینه‌های تمام‌شده', 'مشتریان');

  MatrixSizes: array [0 .. 19] of string = ('2*3', '3*4', '6*4', '6*9', '9*12',
    '10*15', '13*18', '16*21', '20*25', '20*30', '24*30', '30*40', '35*70',
    '40*60', '50*60', '50*70', '60*90', '70*100', '90*120', '110*170');

  CostSizes: array [0 .. 19] of string = ('9*12', '10*15', '13*18', '16*21',
    '20*25', '20*30', '24*30', '30*30', '30*40', '30*45', '30*60', '35*70',
    '40*50', '40*60', '50*60', '50*70', '50*100', '60*90', '70*100', '90*120');

  CostPrefixes: array [1 .. 3] of string = ('p', 's', 'l');

  PriceXlsPath = 'F:\Program sabz\Price.xls'; { پیش‌فرض؛ در SabzPrice.ini قابل تغییر }
  SettingsFile = 'SabzPrice.ini';
  PriceLogFile = 'PriceLog.csv';
  CostInfoText =
    'کد p = چاپ، s = شاسی، l = لمینت. علامت «—» یعنی کدی ثبت نشده است.';

function Blend(C1, C2: TColor; A: Double): TColor;
begin
  Result := RGB(Round(GetRValue(C1) + (GetRValue(C2) - GetRValue(C1)) * A),
    Round(GetGValue(C1) + (GetGValue(C2) - GetGValue(C1)) * A),
    Round(GetBValue(C1) + (GetBValue(C2) - GetBValue(C1)) * A));
end;

procedure TForm1.BtnEnter(Sender: TObject);
begin
  if Sender is TPanel then
    TPanel(Sender).Color := Blend(TColor(TPanel(Sender).Tag), clWhite, 0.16);
end;

procedure TForm1.BtnLeave(Sender: TObject);
begin
  if Sender is TPanel then
    TPanel(Sender).Color := TColor(TPanel(Sender).Tag);
end;

function MakePanel(AParent: TWinControl; AColor: TColor): TPanel;
begin
  Result := TPanel.Create(AParent);
  Result.Parent := AParent;
  Result.BevelOuter := bvNone;
  Result.Color := AColor;
  Result.ParentBackground := False;
  Result.StyleElements := [seFont, seBorder];
  Result.Caption := '';
  Result.DoubleBuffered := True;
end;

function MakeLabel(AParent: TWinControl; const ACap: string; AColor: TColor;
  ASize: Integer; ABold: Boolean = False): TLabel;
begin
  Result := TLabel.Create(AParent);
  Result.Parent := AParent;
  Result.Caption := ACap;
  Result.Font.Name := 'Tahoma';
  Result.Font.Size := ASize;
  Result.Font.Color := AColor;
  if ABold then
    Result.Font.Style := [fsBold];
  Result.Transparent := True;
  Result.BiDiMode := bdRightToLeft;
  Result.Alignment := taRightJustify;
end;

function MakeButton(AParent: TWinControl; const ACap: string; ABg, AFg: TColor;
  AOnClick: TNotifyEvent): TPanel;
begin
  Result := TPanel.Create(AParent);
  Result.Parent := AParent;
  Result.BevelOuter := bvNone;
  Result.Color := ABg;
  Result.ParentBackground := False;
  Result.StyleElements := [seFont, seBorder];
  Result.Caption := ACap;
  Result.Font.Name := 'Tahoma';
  Result.Font.Size := 9;
  Result.Font.Color := AFg;
  Result.Font.Style := [fsBold];
  Result.Alignment := taCenter;
  Result.VerticalAlignment := taVerticalCenter;
  Result.Cursor := crHandPoint;
  Result.BiDiMode := bdRightToLeft;
  Result.Tag := ABg;
  Result.OnClick := AOnClick;
  Result.OnMouseEnter := Form1.BtnEnter;
  Result.OnMouseLeave := Form1.BtnLeave;
end;

function MakeEdit(AParent: TWinControl; AOnChange: TNotifyEvent = nil): TEdit;
begin
  Result := TEdit.Create(AParent);
  Result.Parent := AParent;
  Result.Font.Name := 'Tahoma';
  Result.Font.Size := 10;
  Result.Color := clWhite;
  Result.BiDiMode := bdRightToLeft;
  Result.Alignment := taRightJustify;
  Result.BevelKind := bkFlat;
  if Assigned(AOnChange) then
    Result.OnChange := AOnChange;
end;

function MakeGrid(AParent: TWinControl): TStringGrid;
begin
  Result := TStringGrid.Create(AParent);
  Result.Parent := AParent;
  Result.BevelOuter := bvNone;
  Result.BorderStyle := bsNone;
  Result.Color := clWhite;
  Result.FixedColor := RGB(232, 236, 240);
  Result.Font.Name := 'Tahoma';
  Result.Font.Size := 9;
  Result.Font.Color := C_TEXT;
  Result.DefaultRowHeight := 28;
  Result.RowCount := 2;
  Result.ColCount := 2;
  Result.FixedRows := 1;
  Result.FixedCols := 0;
  Result.BiDiMode := bdLeftToRight;
  Result.DoubleBuffered := True;
  Result.Options := [goVertLine, goHorzLine, goFixedVertLine, goFixedHorzLine,
    goThumbTracking];
  Result.ScrollBars := ssBoth;
  { رسم سلول‌ها کاملاً توسط رویدادهای OnDrawCell فرم انجام می‌شود }
  Result.DefaultDrawing := False;
  Result.DrawingStyle := gdsClassic;
end;

procedure SetupColumns(G: TStringGrid; const Titles: array of string;
  const Widths: array of Integer);
var
  i: Integer;
begin
  G.ColCount := Length(Titles);
  for i := 0 to High(Titles) do
  begin
    G.Cells[i, 0] := Titles[i];
    G.ColWidths[i] := Widths[i];
  end;
  G.RowHeights[0] := 32;
end;

function NewCard(AOwner: TComponent; const ATitle: string): TPanel;
var
  strip: TPanel;
  lbl: TLabel;
begin
  Result := TPanel.Create(AOwner);
  Result.BevelOuter := bvNone;
  Result.Color := C_CARD;
  Result.ParentBackground := False;
  Result.StyleElements := [seFont, seBorder];
  Result.Caption := '';
  Result.DoubleBuffered := True;

  strip := TPanel.Create(Result);
  strip.Parent := Result;
  strip.BevelOuter := bvNone;
  strip.Color := C_ACCENT;
  strip.ParentBackground := False;
  strip.StyleElements := [seFont, seBorder];
  strip.Caption := '';
  strip.Align := alTop;
  strip.Height := 4;

  if ATitle <> '' then
  begin
    lbl := MakeLabel(Result, ATitle, C_TEXT, 11, True);
    lbl.AutoSize := False;
    lbl.Align := alTop;
    lbl.AlignWithMargins := True;
    lbl.Margins.SetBounds(12, 8, 12, 4);
    lbl.Height := 24;
    lbl.Layout := tlCenter;
  end;
end;

procedure DrawGridText(G: TStringGrid; ACol, ARow: Integer; const Rect: TRect;
  const S: string; AAlign: TAlignment);
var
  r: TRect;
  flags: Cardinal;
begin
  r := Rect;
  InflateRect(r, -6, 0);
  flags := DT_SINGLELINE or DT_VCENTER or DT_END_ELLIPSIS or DT_NOPREFIX;
  if G.BiDiMode = bdRightToLeft then
    flags := flags or DT_RTLREADING;
  case AAlign of
    taCenter:
      flags := flags or DT_CENTER;
    taRightJustify:
      flags := flags or DT_RIGHT;
  else
    flags := flags or DT_LEFT;
  end;
  Winapi.Windows.DrawText(G.Canvas.Handle, PChar(S), Length(S), r, flags);
end;

procedure DrawGridFrame(G: TStringGrid; const Rect: TRect);
begin
  G.Canvas.Pen.Color := C_BORDER;
  G.Canvas.Pen.Width := 1;
  G.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
  G.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);
  G.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
  G.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);
end;

function GridColAlign(G: TStringGrid; ACol: Integer): TAlignment;
begin
  Result := taRightJustify;
  if G = Form1.FGridMatrix then
  begin
    if ACol = 0 then
      Result := taCenter;
  end
  else if G = Form1.FGridPrice then
  begin
    if (ACol = 0) or (ACol = 4) then
      Result := taCenter;
  end
  else if G = Form1.FGridCart then
  begin
    if ACol = 1 then
      Result := taCenter;
  end
  else if G = Form1.FGridCost then
  begin
    if ACol = 0 then
      Result := taCenter;
  end
  else if G = Form1.FGridPeople then
  begin
    if ACol in [1, 5, 6] then
      Result := taCenter;
  end
  else if G = Form1.FGridLog then
  begin
    if ACol in [0, 1, 2, 5] then
      Result := taCenter;
  end;
end;

{ رنگ پس‌زمینه و قلم استاندارد سلول (سرستون، انتخاب، ردیف‌های راه‌راه) }
procedure PrepareCell(G: TStringGrid; ARow: Integer; State: TGridDrawState);
begin
  G.Canvas.Font.Style := [];
  G.Canvas.Font.Color := C_TEXT;
  if gdFixed in State then
  begin
    G.Canvas.Brush.Color := C_NAVY;
    G.Canvas.Font.Color := clWhite;
    G.Canvas.Font.Style := [fsBold];
  end
  else if gdSelected in State then
    G.Canvas.Brush.Color := C_SEL
  else if Odd(ARow) then
    G.Canvas.Brush.Color := C_ALT
  else
    G.Canvas.Brush.Color := clWhite;
end;

procedure FinishCell(G: TStringGrid; ACol, ARow: Integer; const Rect: TRect);
var
  al: TAlignment;
begin
  G.Canvas.FillRect(Rect);
  DrawGridFrame(G, Rect);
  if ARow < G.FixedRows then
    al := taCenter
  else
    al := GridColAlign(G, ACol);
  DrawGridText(G, ACol, ARow, Rect, G.Cells[ACol, ARow], al);
end;

function AppDir: string;
begin
  Result := ExtractFilePath(Application.ExeName);
end;

{ ------------------------------- TCartLine ------------------------------- }

function TCartLine.Total: Int64;
begin
  Result := Int64(Qty) * UnitPrice;
end;

{ -------------------------------- Form ----------------------------------- }

procedure TForm1.FormCreate(Sender: TObject);
begin
  Caption := 'سامانه قیمت استودیو سبز';
  ClientWidth := 1280;
  ClientHeight := 800;
  Constraints.MinWidth := 1080;
  Constraints.MinHeight := 680;
  Position := poScreenCenter;
  Color := C_BG;
  Font.Name := 'Tahoma';
  Font.Size := 9;
  BiDiMode := bdRightToLeft;
  DoubleBuffered := True;
  KeyPreview := True;

  OnKeyDown := FormKeyDown;
  OnCloseQuery := FormCloseQuery;

  FData := TSabzData.Create;
  FPrices := TObjectList<TPriceItem>.Create(True);
  FIndex := TDictionary<string, TPriceItem>.Create;
  FCart := TObjectList<TCartLine>.Create(True);
  FOriginal := TDictionary<string, Int64>.Create;
  FPctBackup := TDictionary<string, Int64>.Create;
  FCostDraft := TDictionary<string, Int64>.Create;
  FPeopleDb := TObjectList<TPerson>.Create(True);
  FAllPeople := TObjectList<TPerson>.Create(True);
  FPeopleView := TList<TPerson>.Create;
  FPeopleDist := TList<Integer>.Create;
  FCurrentPage := -1;

  FPeopleTimer := TTimer.Create(Self);
  FPeopleTimer.Enabled := False;
  FPeopleTimer.Interval := 350;
  FPeopleTimer.OnTimer := PeopleTimerTimer;

  LoadSettings;

  BuildHeader;
  BuildFooter;
  BuildNav;
  BuildDashboard;
  BuildMatrixPage;
  BuildPricePage;
  BuildCalcPage;
  BuildCostPage;
  BuildPeoplePage;

  ShowPage(StrToIntDef(ParamStr(1), 0));

  if FData.Connect then
    ReloadPrices
  else
  begin
    UpdateConnStatus;
    SetStatus('اتصال به پایگاه داده برقرار نشد. دکمه «اتصال مجدد» را بزنید.');
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FPeopleTimer.Enabled := False;
  FPeopleDist.Free;
  FPeopleView.Free;
  FAllPeople.Free;
  FPeopleDb.Free;
  FIndex.Free;
  FCostDraft.Free;
  FPctBackup.Free;
  FOriginal.Free;
  FCart.Free;
  FPrices.Free;
  FData.Free;
end;

procedure TForm1.BuildHeader;
var
  hdr, brand, info, line: TPanel;
begin
  hdr := TPanel.Create(Self);
  hdr.Parent := Self;
  hdr.Align := alTop;
  hdr.Height := 74;
  hdr.BevelOuter := bvNone;
  hdr.Color := C_NAVY;
  hdr.ParentBackground := False;
  hdr.StyleElements := [seFont, seBorder];
  hdr.Caption := '';
  hdr.DoubleBuffered := True;

  brand := MakePanel(hdr, C_NAVY);
  brand.Align := alRight;
  brand.Width := 360;

  line := MakePanel(brand, C_GOLD);
  line.Align := alRight;
  line.Width := 5;

  with MakeLabel(brand, 'استودیو عکاسی سبز', clWhite, 16, True) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(10, 12, 16, 0);
    Height := 30;
  end;
  with MakeLabel(brand, 'سامانه مدیریت قیمت و محاسبه سفارش', C_ACCENT_L, 9) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(10, 0, 18, 0);
    Height := 20;
  end;

  info := MakePanel(hdr, C_NAVY);
  info.Align := alLeft;
  info.Width := 400;

  FLblConn := MakeLabel(info, 'در حال اتصال...', C_ACCENT_L, 9, True);
  FLblConn.Left := 14;
  FLblConn.Top := 16;
  FLblConn.AutoSize := True;

  with MakeLabel(info, 'تاریخ: ' + ToPersianDigits(JalaliToday), clWhite, 9) do
  begin
    Left := 14;
    Top := 40;
    AutoSize := True;
  end;

  FBtnReconnect := MakeButton(info, 'اتصال مجدد', C_GOLD, C_NAVY,
    BtnReconnectClick);
  FBtnReconnect.Left := 270;
  FBtnReconnect.Top := 20;
  FBtnReconnect.Width := 115;
  FBtnReconnect.Height := 34;
  FBtnReconnect.Visible := False;
end;

procedure TForm1.BuildFooter;
var
  ft: TPanel;
begin
  ft := MakePanel(Self, C_CARD);
  ft.Align := alBottom;
  ft.Height := 30;

  FStatus := MakeLabel(ft, 'آماده', C_MUTED, 9);
  FStatus.Align := alClient;
  FStatus.AlignWithMargins := True;
  FStatus.Margins.SetBounds(14, 0, 14, 0);
  FStatus.Layout := tlCenter;
  FStatus.AutoSize := False;
  FStatus.Alignment := taRightJustify;

  with MakeLabel(ft,
    'F1..F6 صفحات  |  Ctrl+S ذخیره  |  Ctrl+F جستجو  |  Ctrl+R بارگذاری مجدد',
    C_MUTED, 8) do
  begin
    Align := alLeft;
    AlignWithMargins := True;
    Margins.SetBounds(8, 0, 8, 0);
    Layout := tlCenter;
    AutoSize := False;
    Width := 420;
    Alignment := taLeftJustify;
    BiDiMode := bdLeftToRight;
  end;

  with MakeLabel(ft, 'نسخه ۳٫۰', C_MUTED, 8) do
  begin
    Align := alLeft;
    AlignWithMargins := True;
    Margins.SetBounds(14, 0, 8, 0);
    Layout := tlCenter;
    AutoSize := False;
    Width := 70;
    Alignment := taLeftJustify;
  end;
end;

procedure TForm1.BuildNav;
var
  nav: TPanel;
  i: Integer;
  b: TPanel;
begin
  nav := MakePanel(Self, C_NAVY);
  nav.Align := alRight;
  nav.Width := 208;

  with MakeLabel(nav, 'منو', C_MUTED, 8) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(16, 16, 16, 6);
    Height := 18;
  end;

  for i := 0 to High(NavTitles) do
  begin
    b := TPanel.Create(nav);
    b.Parent := nav;
    b.Align := alTop;
    b.Height := 46;
    b.BevelOuter := bvNone;
    b.Color := C_NAVY;
    b.ParentBackground := False;
    b.StyleElements := [seFont, seBorder];
    b.Caption := '  ' + NavTitles[i];
    b.Font.Name := 'Tahoma';
    b.Font.Size := 10;
    b.Font.Color := clWhite;
    b.Alignment := taRightJustify;
    b.VerticalAlignment := taVerticalCenter;
    b.Cursor := crHandPoint;
    b.BiDiMode := bdRightToLeft;
    b.Tag := i;
    b.OnClick := NavClick;
    b.OnMouseEnter := NavEnter;
    b.OnMouseLeave := NavLeave;
    b.AlignWithMargins := True;
    b.Margins.SetBounds(10, 2, 10, 2);
    b.Hint := 'میان‌بر: F' + IntToStr(i + 1);
    b.ShowHint := True;
    FNav[i] := b;
  end;
end;

procedure TForm1.BuildDashboard;
var
  page, stats, body, card, content: TPanel;
  guide: TLabel;
  i: Integer;
  caps: array [0 .. 3] of string;
begin
  page := MakePanel(Self, C_BG);
  page.Parent := Self;
  page.Align := alClient;
  page.Visible := False;
  FPages[0] := page;

  stats := MakePanel(page, C_BG);
  stats.Align := alTop;
  stats.Height := 132;
  stats.OnResize := LayoutStats;

  caps[0] := 'ردیف‌های قیمت';
  caps[1] := 'مشتریان';
  caps[2] := 'تاریخ امروز';
  caps[3] := 'وضعیت اتصال';
  for i := 0 to 3 do
  begin
    card := NewCard(stats, '');
    card.Parent := stats;
    card.AlignWithMargins := True;
    card.Margins.SetBounds(6, 6, 6, 6);
    card.Align := alRight;
    card.Width := 250;
    FStatCards[i] := card;
    FLblStat[i] := MakeLabel(card, '—', C_ACCENT, 20, True);
    FLblStat[i].Align := alClient;
    FLblStat[i].AutoSize := False;
    FLblStat[i].Layout := tlCenter;
    FLblStat[i].Alignment := taRightJustify;
    with MakeLabel(card, caps[i], C_MUTED, 9) do
    begin
      Align := alBottom;
      AlignWithMargins := True;
      Margins.SetBounds(12, 0, 12, 10);
      Height := 20;
    end;
  end;

  body := MakePanel(page, C_BG);
  body.Align := alClient;

  card := NewCard(body, 'دسترسی سریع');
  card.Parent := body;
  card.AlignWithMargins := True;
  card.Margins.SetBounds(6, 6, 6, 6);
  card.Align := alRight;
  card.Width := 340;
  content := card;
  with MakeButton(content, 'ویرایش قیمت‌ها', C_ACCENT, clWhite, BtnGoEditClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'محاسبه سفارش', C_NAVY, clWhite, BtnGoCalcClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'جدول قیمت', C_GOLD, C_NAVY, BtnGoMatrixClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'هزینه‌های تمام‌شده', C_GOLD, C_NAVY,
    BtnGoCostClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'خواندن از اکسل (آپدیت دیتابیس)', C_NAVY, clWhite,
    BtnFromExcelClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'ذخیره در اکسل (Price.xls)', C_ACCENT, clWhite,
    BtnToExcelClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'مشتریان (جستجوی هوشمند)', C_NAVY, clWhite,
    BtnGoPeopleClick) do
  begin
    Align := alTop;
    AlignWithMargins := True;
    Margins.SetBounds(0, 0, 0, 8);
    Height := 44;
  end;
  with MakeButton(content, 'بارگذاری مجدد قیمت‌ها', C_BORDER, C_TEXT,
    BtnReconnectClick) do
  begin
    Align := alTop;
    Height := 44;
  end;

  { ---- تاریخچه آخرین تغییرات قیمت ---- }
  card := NewCard(body, 'آخرین تغییرات قیمت (تاریخچه)');
  card.Parent := body;
  card.AlignWithMargins := True;
  card.Margins.SetBounds(6, 6, 6, 6);
  card.Align := alLeft;
  card.Width := 470;
  FGridLog := MakeGrid(card);
  FGridLog.Align := alClient;
  FGridLog.Options := FGridLog.Options + [goRowSelect];
  SetupColumns(FGridLog, ['تاریخ', 'ساعت', 'کد', 'قبلی', 'جدید', 'تغییر'],
    [78, 48, 78, 88, 88, 62]);
  FGridLog.OnDrawCell := GridLogDrawCell;

  card := NewCard(body, 'راهنمای کدها و معنی قیمت‌ها');
  card.Parent := body;
  card.AlignWithMargins := True;
  card.Margins.SetBounds(6, 6, 6, 6);
  card.Align := alClient;
  content := card;
  guide := MakeLabel(content,
    '  • کدهای پایانیافته با «n» → عکس هنری، فقط چاپ (عکس جدید)' + sLineBreak +
    '  • سایز ۹×۱۲ به بالا بدون پسوند → عکس هنری، چاپ مجدد' + sLineBreak +
    '  • کدهای پایانیافته با «_» → قیمت شاسی و لمینت تنها' + sLineBreak +
    '  • کدهای دارای «f» → هزینه هر چهره اضافه برای روتوش' + sLineBreak +
    '  • کدهای دارای «i» → آلبوم ایتالیایی عروس و داماد' + sLineBreak +
    '  • کدهای دارای «a» → عکس مراسم عروس و داماد' + sLineBreak +
    '  • شروع با «s» → قیمت تمام‌شده شاسی برای ما' + sLineBreak +
    '  • شروع با «p» → قیمت تمام‌شده چاپ برای ما' + sLineBreak +
    '  • شروع با «f_» → فقط فایل' + sLineBreak +
    '  • بقیه کدها پرسنلی یا طبق نام مشخص هستند', C_TEXT, 10);
  guide.Align := alClient;
  guide.AutoSize := False;
  guide.WordWrap := True;
  guide.Layout := tlTop;
  guide.Alignment := taRightJustify;
end;

procedure TForm1.BuildMatrixPage;
var
  page, card, content, hdr: TPanel;
begin
  page := MakePanel(Self, C_BG);
  page.Parent := Self;
  page.Align := alClient;
  page.Visible := False;
  FPages[1] := page;

  hdr := MakePanel(page, C_BG);
  hdr.Align := alTop;
  hdr.Height := 46;
  with MakeLabel(hdr, 'جدول قیمت عکس هنری (خوانا)', C_TEXT, 13, True) do
  begin
    Align := alClient;
    AutoSize := False;
    Layout := tlCenter;
  end;

  card := NewCard(page, '');
  card.Parent := page;
  card.Align := alClient;
  content := card;

  FGridMatrix := MakeGrid(content);
  FGridMatrix.Align := alClient;
  FGridMatrix.Options := FGridMatrix.Options + [goRowSelect];
  SetupColumns(FGridMatrix,
    ['سایز', 'چاپ مجدد', 'عکس جدید', 'شاسی', 'مجدد+شاسی', 'جدید+شاسی',
    'چهره اضافه', 'سود چاپ مجدد'],
    [100, 130, 130, 130, 140, 140, 120, 130]);
  FGridMatrix.RowCount := 1 + Length(MatrixSizes);
  FGridMatrix.OnDrawCell := GridMatrixDrawCell;
  FGridMatrix.Hint := 'سود چاپ مجدد = قیمت چاپ مجدد − (هزینه چاپ + لمینت)';
  FGridMatrix.ShowHint := True;
end;

procedure TForm1.BuildPricePage;
var
  page, card, content, toolbar, pctTools, hdr: TPanel;
begin
  page := MakePanel(Self, C_BG);
  page.Parent := Self;
  page.Align := alClient;
  page.Visible := False;
  FPages[2] := page;

  hdr := MakePanel(page, C_BG);
  hdr.Align := alTop;
  hdr.Height := 46;
  with MakeLabel(hdr, 'ویرایش قیمت‌ها (جدول Price)', C_TEXT, 13, True) do
  begin
    Align := alClient;
    AutoSize := False;
    Layout := tlCenter;
  end;

  card := NewCard(page, '');
  card.Parent := page;
  card.Align := alClient;
  content := card;

  toolbar := MakePanel(content, C_CARD);
  toolbar.Align := alTop;
  toolbar.Height := 52;

  FEdPriceSearch := MakeEdit(toolbar, EdPriceSearchChange);
  FEdPriceSearch.Left := 8;
  FEdPriceSearch.Top := 10;
  FEdPriceSearch.Width := 260;
  FEdPriceSearch.Height := 32;
  FEdPriceSearch.TextHint := 'جستجو در شرح یا کد (مثلاً 10x15)...';

  FCbCat := TComboBox.Create(toolbar);
  FCbCat.Parent := toolbar;
  FCbCat.Left := 276;
  FCbCat.Top := 10;
  FCbCat.Width := 220;
  FCbCat.Height := 32;
  FCbCat.Style := csDropDownList;
  FCbCat.Font.Name := 'Tahoma';
  FCbCat.Font.Size := 10;
  FCbCat.BiDiMode := bdRightToLeft;
  FCbCat.OnChange := CbCatChange;
  FCbCat.Items.Add('همه دسته‌ها');
  FCbCat.Items.Add(PriceCategoryNames[pcArtReprint]);
  FCbCat.Items.Add(PriceCategoryNames[pcArtNew]);
  FCbCat.Items.Add(PriceCategoryNames[pcFrame]);
  FCbCat.Items.Add(PriceCategoryNames[pcFace]);
  FCbCat.Items.Add(PriceCategoryNames[pcItalianAlbum]);
  FCbCat.Items.Add(PriceCategoryNames[pcCeremony]);
  FCbCat.Items.Add(PriceCategoryNames[pcPrintCost]);
  FCbCat.Items.Add(PriceCategoryNames[pcFrameCost]);
  FCbCat.Items.Add(PriceCategoryNames[pcLaminate]);
  FCbCat.Items.Add(PriceCategoryNames[pcFile]);
  FCbCat.Items.Add(PriceCategoryNames[pcOther]);
  FCbCat.ItemIndex := 0;

  FBtnSave := MakeButton(toolbar, 'ذخیره تغییرات', C_ACCENT, clWhite,
    BtnSaveClick);
  FBtnSave.Left := 508;
  FBtnSave.Top := 10;
  FBtnSave.Width := 140;
  FBtnSave.Height := 32;

  FLblChanges := MakeLabel(toolbar, '', C_DANGER, 9, True);
  FLblChanges.Left := 660;
  FLblChanges.Top := 18;
  FLblChanges.AutoSize := True;

  { ---- نوار افزایش/کاهش درصدی با گرد کردن ---- }
  pctTools := MakePanel(content, C_ALT);
  pctTools.Align := alTop;
  pctTools.Height := 56;

  with MakeLabel(pctTools, 'افزایش/کاهش درصدی:', C_TEXT, 9, True) do
  begin
    Left := 10;
    Top := 20;
    AutoSize := True;
  end;

  FEdPercent := MakeEdit(pctTools);
  FEdPercent.Left := 140;
  FEdPercent.Top := 12;
  FEdPercent.Width := 80;
  FEdPercent.Height := 32;
  FEdPercent.Text := '20';
  FEdPercent.TextHint := 'مثلاً 20';

  with MakeLabel(pctTools, '٪', C_MUTED, 10, True) do
  begin
    Left := 224;
    Top := 20;
    AutoSize := True;
  end;

  with MakeLabel(pctTools, 'گرد کردن:', C_MUTED, 9, True) do
  begin
    Left := 254;
    Top := 20;
    AutoSize := True;
  end;

  FCbRound := TComboBox.Create(pctTools);
  FCbRound.Parent := pctTools;
  FCbRound.Left := 316;
  FCbRound.Top := 12;
  FCbRound.Width := 150;
  FCbRound.Height := 32;
  FCbRound.Style := csDropDownList;
  FCbRound.Font.Name := 'Tahoma';
  FCbRound.Font.Size := 10;
  FCbRound.BiDiMode := bdRightToLeft;
  FCbRound.Items.Add('بدون گرد کردن');
  FCbRound.Items.Add('۱،۰۰۰ تومان');
  FCbRound.Items.Add('۵،۰۰۰ تومان');
  FCbRound.Items.Add('۱۰،۰۰۰ تومان');
  FCbRound.Items.Add('۵۰،۰۰۰ تومان');
  FCbRound.Items.Add('۱۰۰،۰۰۰ تومان');
  FCbRound.ItemIndex := 3;

  FCbRoundMode := TComboBox.Create(pctTools);
  FCbRoundMode.Parent := pctTools;
  FCbRoundMode.Left := 472;
  FCbRoundMode.Top := 12;
  FCbRoundMode.Width := 120;
  FCbRoundMode.Height := 32;
  FCbRoundMode.Style := csDropDownList;
  FCbRoundMode.Font.Name := 'Tahoma';
  FCbRoundMode.Font.Size := 10;
  FCbRoundMode.BiDiMode := bdRightToLeft;
  FCbRoundMode.Items.Add('نزدیک‌ترین');
  FCbRoundMode.Items.Add('به بالا');
  FCbRoundMode.Items.Add('به پایین');
  FCbRoundMode.ItemIndex := 0;

  FBtnApplyPct := MakeButton(pctTools, 'اعمال درصد', C_ACCENT, clWhite,
    ApplyPctClick);
  FBtnApplyPct.Left := 598;
  FBtnApplyPct.Top := 12;
  FBtnApplyPct.Width := 130;
  FBtnApplyPct.Height := 32;

  FBtnUndoPct := MakeButton(pctTools, 'بازگرداندن', C_BORDER, C_TEXT,
    UndoPctClick);
  FBtnUndoPct.Left := 734;
  FBtnUndoPct.Top := 12;
  FBtnUndoPct.Width := 120;
  FBtnUndoPct.Height := 32;

  FLblPctInfo := MakeLabel(pctTools, 'روی ردیف‌های فیلترشده اعمال می‌شود.', C_MUTED, 8);
  FLblPctInfo.Left := 862;
  FLblPctInfo.Top := 20;
  FLblPctInfo.AutoSize := True;

  FGridPrice := MakeGrid(content);
  FGridPrice.Align := alClient;
  FGridPrice.Options := FGridPrice.Options + [goEditing];
  SetupColumns(FGridPrice,
    ['کد', 'شرح', 'دسته', 'قیمت (تومان)', 'تعداد', 'وضعیت / مقدار قبلی'],
    [85, 290, 180, 120, 50, 210]);
  FGridPrice.OnDrawCell := GridPriceDrawCell;
  FGridPrice.OnSelectCell := GridPriceSelectCell;
  FGridPrice.OnSetEditText := GridPriceSetEditText;
end;

procedure TForm1.BuildCalcPage;
var
  page, outer, catCard, catTool, cartCard, cartTool, totals, custRow: TPanel;
begin
  page := MakePanel(Self, C_BG);
  page.Parent := Self;
  page.Align := alClient;
  page.Visible := False;
  FPages[3] := page;

  outer := MakePanel(page, C_BG);
  outer.Align := alClient;
  outer.OnResize := LayoutCalc;
  FCalcOuter := outer;

  { --- catalog (سمت راست) --- }
  catCard := NewCard(outer, 'لیست قیمت‌ها (برای افزودن دوبار کلیک کنید)');
  catCard.Parent := outer;
  catCard.AlignWithMargins := True;
  catCard.Margins.SetBounds(6, 6, 6, 6);
  catCard.Align := alRight;
  catCard.Width := 500;
  FCalcCatCard := catCard;

  catTool := MakePanel(catCard, C_CARD);
  catTool.Align := alTop;
  catTool.Height := 46;
  FEdCatSearch := MakeEdit(catTool, EdCatSearchChange);
  FEdCatSearch.Left := 8;
  FEdCatSearch.Top := 7;
  FEdCatSearch.Width := 240;
  FEdCatSearch.Height := 32;
  FEdCatSearch.TextHint := 'جستجو... (Enter = افزودن)';
  FEdCatSearch.OnKeyPress := EdSearchKeyPress;
  with MakeButton(catTool, 'افزودن', C_ACCENT, clWhite, BtnAddClick) do
  begin
    Left := 256;
    Top := 7;
    Width := 100;
    Height := 32;
  end;

  FGridCatalog := MakeGrid(catCard);
  FGridCatalog.Align := alClient;
  FGridCatalog.Options := FGridCatalog.Options + [goRowSelect];
  SetupColumns(FGridCatalog, ['شرح', 'قیمت (تومان)'], [300, 140]);
  FGridCatalog.OnDblClick := GridCatDblClick;
  FGridCatalog.OnDrawCell := GridCatDrawCell;

  { --- cart (سمت چپ) --- }
  cartCard := NewCard(outer, 'سبد سفارش');
  cartCard.Parent := outer;
  cartCard.AlignWithMargins := True;
  cartCard.Margins.SetBounds(6, 6, 6, 6);
  cartCard.Align := alClient;

  custRow := MakePanel(cartCard, C_CARD);
  custRow.Align := alTop;
  custRow.Height := 40;
  with MakeLabel(custRow, 'نام مشتری:', C_MUTED, 9, True) do
  begin
    Left := 12;
    Top := 12;
    AutoSize := True;
  end;
  FEdCustomer := MakeEdit(custRow);
  FEdCustomer.Left := 100;
  FEdCustomer.Top := 6;
  FEdCustomer.Width := 300;
  FEdCustomer.Height := 30;
  FEdCustomer.TextHint := 'نام مشتری (اختیاری)';
  FEdCustomer.OnChange := EdCustomerChange;
  with MakeButton(custRow, 'انتخاب از مشتریان', C_GOLD, C_NAVY,
    BtnGoPeopleClick) do
  begin
    Left := 408;
    Top := 5;
    Width := 130;
    Height := 30;
  end;

  cartTool := MakePanel(cartCard, C_CARD);
  cartTool.Align := alTop;
  cartTool.Height := 126;

  with MakeButton(cartTool, 'حذف انتخاب', C_DANGER, clWhite,
    BtnRemoveClick) do
  begin
    Left := 8;
    Top := 7;
    Width := 100;
    Height := 32;
  end;
  with MakeButton(cartTool, 'پاک کردن', C_BORDER, C_TEXT,
    BtnClearCartClick) do
  begin
    Left := 114;
    Top := 7;
    Width := 100;
    Height := 32;
  end;
  with MakeButton(cartTool, 'چاپ فاکتور A5', C_NAVY, clWhite,
    BtnPrintInvoiceClick) do
  begin
    Left := 220;
    Top := 7;
    Width := 122;
    Height := 32;
  end;
  with MakeButton(cartTool, 'ذخیره فاکتور', C_SUCCESS, clWhite,
    BtnSaveInvoiceClick) do
  begin
    Left := 348;
    Top := 7;
    Width := 112;
    Height := 32;
  end;
  with MakeButton(cartTool, 'کپی', C_BORDER, C_TEXT, BtnCopyInvoiceClick) do
  begin
    Left := 466;
    Top := 7;
    Width := 66;
    Height := 32;
  end;

  with MakeButton(cartTool, 'شاسی و لمینت با هم', C_ACCENT, clWhite,
    BtnSvcBothClick) do
  begin
    Left := 8;
    Top := 48;
    Width := 148;
    Height := 32;
  end;
  with MakeButton(cartTool, 'شاسی جدا', C_GOLD, C_TEXT, BtnSvcFrameClick) do
  begin
    Left := 162;
    Top := 48;
    Width := 88;
    Height := 32;
  end;
  with MakeButton(cartTool, 'لمینت جدا', C_GOLD, C_TEXT, BtnSvcLamClick) do
  begin
    Left := 256;
    Top := 48;
    Width := 88;
    Height := 32;
  end;
  with MakeButton(cartTool, 'باز کردن فاکتورها', C_NAVY, clWhite,
    BtnOpenFacturesClick) do
  begin
    Left := 350;
    Top := 48;
    Width := 140;
    Height := 32;
  end;

  with MakeLabel(cartTool, 'تاریخ فاکتور:', C_MUTED, 9, True) do
  begin
    Left := 8;
    Top := 92;
    AutoSize := True;
  end;
  FEdInvDate := MakeEdit(cartTool);
  FEdInvDate.Left := 96;
  FEdInvDate.Top := 87;
  FEdInvDate.Width := 118;
  FEdInvDate.Height := 30;
  FEdInvDate.Text := ToPersianDigits(JalaliToday);
  FEdInvDate.TextHint := 'مثلاً ۱۴۰۵/۰۷/۰۸';
  FEdInvDate.Hint := 'تاریخ روز فاکتور؛ قابل ویرایش است و روی هر دو نوع فاکتور می‌آید';
  FEdInvDate.ShowHint := True;
  FEdInvDate.OnEnter := EdInvDateEnter;
  FEdInvDate.OnClick := EdInvDateEnter;
  with MakeButton(cartTool, 'امروز', C_BORDER, C_TEXT, BtnInvTodayClick) do
  begin
    Left := 222;
    Top := 88;
    Width := 74;
    Height := 30;
  end;

  totals := MakePanel(cartCard, C_ALT);
  totals.Align := alBottom;
  totals.Height := 176;

  with MakeLabel(totals, 'جمع فروش:', C_TEXT, 10, True) do
  begin
    Left := 14;
    Top := 12;
    AutoSize := True;
  end;
  FLblSubtotal := MakeLabel(totals, '0 تومان', C_TEXT, 11, True);
  FLblSubtotal.Left := 200;
  FLblSubtotal.Top := 10;
  FLblSubtotal.AutoSize := True;

  with MakeLabel(totals, 'جمع هزینه ما:', C_TEXT, 10, True) do
  begin
    Left := 14;
    Top := 42;
    AutoSize := True;
  end;
  FLblCost := MakeLabel(totals, '0 تومان', C_MUTED, 11, True);
  FLblCost.Left := 200;
  FLblCost.Top := 40;
  FLblCost.AutoSize := True;

  with MakeLabel(totals, 'تخفیف (تومان یا ٪):', C_TEXT, 10, True) do
  begin
    Left := 14;
    Top := 72;
    AutoSize := True;
  end;
  FEdDisc := MakeEdit(totals, EdDiscChange);
  FEdDisc.Left := 200;
  FEdDisc.Top := 68;
  FEdDisc.Width := 160;
  FEdDisc.Height := 30;
  FEdDisc.Text := '0';
  FEdDisc.TextHint := 'مثلاً 50000 یا 10%';
  FEdDisc.Hint := 'مبلغ تخفیف به تومان، یا درصد با علامت ٪ (مثلاً 10%)';
  FEdDisc.ShowHint := True;

  FChkShop := TCheckBox.Create(totals);
  FChkShop.Parent := totals;
  FChkShop.Left := 366;
  FChkShop.Top := 72;
  FChkShop.Width := 150;
  FChkShop.Height := 22;
  FChkShop.Caption := 'فاکتور با نام چاپخانه';
  FChkShop.Font.Name := 'Tahoma';
  FChkShop.Font.Size := 9;
  FChkShop.BiDiMode := bdRightToLeft;
  FChkShop.Hint := 'اگر تیک بخورد، فاکتور به نام چاپخانه (از SabzPrice.ini) صادر می‌شود';
  FChkShop.ShowHint := True;

  with MakeLabel(totals, 'سود خالص:', C_SUCCESS, 11, True) do
  begin
    Left := 14;
    Top := 110;
    AutoSize := True;
  end;
  FLblProfit := MakeLabel(totals, '0 تومان', C_SUCCESS, 12, True);
  FLblProfit.Left := 200;
  FLblProfit.Top := 108;
  FLblProfit.AutoSize := True;

  with MakeLabel(totals, 'قابل پرداخت:', C_ACCENT, 12, True) do
  begin
    Left := 14;
    Top := 142;
    AutoSize := True;
  end;
  FLblPayable := MakeLabel(totals, '0 تومان', C_ACCENT, 14, True);
  FLblPayable.Left := 200;
  FLblPayable.Top := 139;
  FLblPayable.AutoSize := True;

  FGridCart := MakeGrid(cartCard);
  FGridCart.Align := alClient;
  FGridCart.Options := FGridCart.Options + [goEditing];
  SetupColumns(FGridCart,
    ['شرح', 'تعداد', 'قیمت واحد', 'هزینه ما', 'جمع'],
    [160, 55, 95, 95, 100]);
  FGridCart.OnDrawCell := GridCartDrawCell;
  FGridCart.OnSelectCell := GridCartSelectCell;
  FGridCart.OnSetEditText := GridCartSetEditText;
end;

procedure TForm1.BuildCostPage;
var
  page, card, content, toolbar, hdr: TPanel;
begin
  page := MakePanel(Self, C_BG);
  page.Parent := Self;
  page.Align := alClient;
  page.Visible := False;
  FPages[4] := page;

  hdr := MakePanel(page, C_BG);
  hdr.Align := alTop;
  hdr.Height := 46;
  with MakeLabel(hdr, 'هزینه‌های تمام‌شده (چاپ / شاسی / لمینت)', C_TEXT, 13,
    True) do
  begin
    Align := alClient;
    AutoSize := False;
    Layout := tlCenter;
  end;

  card := NewCard(page, '');
  card.Parent := page;
  card.Align := alClient;
  content := card;

  toolbar := MakePanel(content, C_CARD);
  toolbar.Align := alTop;
  toolbar.Height := 52;

  FBtnCostSave := MakeButton(toolbar, 'ذخیره هزینه‌ها', C_ACCENT, clWhite,
    BtnCostSaveClick);
  FBtnCostSave.Left := 12;
  FBtnCostSave.Top := 10;
  FBtnCostSave.Width := 150;
  FBtnCostSave.Height := 32;

  FLblCostInfo := MakeLabel(toolbar, CostInfoText, C_MUTED, 8);
  FLblCostInfo.Left := 174;
  FLblCostInfo.Top := 18;
  FLblCostInfo.AutoSize := True;

  FGridCost := MakeGrid(content);
  FGridCost.Align := alClient;
  FGridCost.Options := FGridCost.Options + [goEditing];
  SetupColumns(FGridCost,
    ['سایز', 'چاپ (تومان)', 'شاسی (تومان)', 'لمینت (تومان)'],
    [130, 170, 170, 170]);
  FGridCost.OnDrawCell := GridCostDrawCell;
  FGridCost.OnSelectCell := GridCostSelectCell;
  FGridCost.OnSetEditText := GridCostSetEditText;
end;

procedure TForm1.FillCostGrid;
var
  i, c: Integer;
  it: TPriceItem;
  code: string;
  v: Int64;
begin
  if FGridCost = nil then
    Exit;
  FUpdating := True;
  try
    FGridCost.RowCount := 1 + Length(CostSizes);
    for i := 0 to High(CostSizes) do
    begin
      FGridCost.Cells[0, i + 1] := CostSizes[i];
      for c := 1 to 3 do
      begin
        code := CostPrefixes[c] + CostSizes[i];
        if FCostDraft.TryGetValue(code, v) then
          FGridCost.Cells[c, i + 1] := FormatMoney(v)
        else
        begin
          it := FindItem(code);
          if it <> nil then
            FGridCost.Cells[c, i + 1] := FormatMoney(it.Price)
          else
            FGridCost.Cells[c, i + 1] := '—';
        end;
      end;
    end;
    if FGridCost.RowCount > 1 then
      FGridCost.Row := 1;
  finally
    FUpdating := False;
  end;
end;

procedure TForm1.GridCostSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  CanSelect := True;
  FGridCost.EditorMode := (ACol >= 1) and (ACol <= 3) and (ARow >= 1);
end;

procedure TForm1.GridCostSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: string);
var
  code: string;
  v: Int64;
  it: TPriceItem;
begin
  if FUpdating then
    Exit;
  if (ACol < 1) or (ACol > 3) or (ARow < 1) then
    Exit;
  code := CostPrefixes[ACol] + FGridCost.Cells[0, ARow];
  v := ParseMoney(Value);
  it := FindItem(code);
  if ((it <> nil) and (v = it.Price)) or ((it = nil) and (v = 0)) then
  begin
    { بدون تغییر واقعی (مثلاً فقط کلیک) }
    FCostDraft.Remove(code);
  end
  else
  begin
    FCostDraft.AddOrSetValue(code, v);
    SetStatus('هزینه «' + code + '» به ' + FormatToman(v) +
      ' تغییر یافت. برای ثبت، «ذخیره هزینه‌ها» را بزنید.');
  end;
  if FCostDraft.Count = 0 then
    FLblCostInfo.Caption := CostInfoText
  else
    FLblCostInfo.Caption := IntToStr(FCostDraft.Count) +
      ' هزینه تغییر کرده (ذخیره نشده)';
end;

procedure TForm1.BtnCostSaveClick(Sender: TObject);
begin
  if FCostDraft.Count = 0 then
  begin
    SetStatus('هزینه تغییریافته‌ای برای ذخیره وجود ندارد.');
    Exit;
  end;
  { ذخیره هزینه‌ها قیمت‌ها را دوباره بارگذاری می‌کند؛ تغییرات صفحه ویرایش
    قیمت نباید بی‌صدا از بین برود }
  if ConfirmPending(False) then
    DoSaveCosts;
end;

function TForm1.DoSaveCosts: Boolean;
var
  pair: TPair<string, Int64>;
  it: TPriceItem;
  ok, fail, ins: Integer;
begin
  ok := 0;
  fail := 0;
  ins := 0;
  Screen.Cursor := crHourGlass;
  try
    for pair in FCostDraft do
    begin
      it := FindItem(pair.Key);
      if it <> nil then
      begin
        if it.Price <> pair.Value then
        begin
          if FData.SavePrice(pair.Key, pair.Value) then
          begin
            LogPriceChange(pair.Key, it.Price, pair.Value, 'هزینه');
            Inc(ok);
          end
          else
            Inc(fail);
        end;
      end
      else
      begin
        if FData.InsertPrice(pair.Key, pair.Value) then
        begin
          LogPriceChange(pair.Key, -1, pair.Value, 'هزینه (جدید)');
          Inc(ins);
        end
        else
          Inc(fail);
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  Result := fail = 0;
  FCostDraft.Clear;
  ReloadPrices;
  FLblCostInfo.Caption := CostInfoText;
  if fail = 0 then
    SetStatus(Format('%d به‌روزرسانی و %d ردیف جدید ذخیره شد.',
      [ok, ins]))
  else
    Application.MessageBox(PChar(IntToStr(fail) + ' ردیف ذخیره نشد.' +
      sLineBreak + FData.LastError), 'خطا در ذخیره', MB_OK or MB_ICONERROR);
end;

procedure TForm1.BtnGoCostClick(Sender: TObject);
begin
  ShowPage(4);
end;

procedure TForm1.BtnFromExcelClick(Sender: TObject);
var
  XL, WB, WS: OleVariant;
  pv: OleVariant;
  r, n, upd, ins: Integer;
  code: string;
  v: Int64;
  it: TPriceItem;
begin
  if not ConfirmPending(True) then
    Exit;
  if not ResolveExcelPath then
    Exit;
  Screen.Cursor := crHourGlass;
  try
    try
      XL := CreateOleObject('Excel.Application');
      try
        XL.Visible := False;
        XL.DisplayAlerts := False;
        WB := XL.Workbooks.Open(FExcelPath);
        try
          WS := WB.Worksheets['Price'];
          n := WS.UsedRange.Rows.Count;
          upd := 0;
          ins := 0;
          for r := 2 to n do
          begin
            code := Trim(VarToStr(WS.Cells[r, 1].Value));
            if code = '' then
              Continue;
            pv := WS.Cells[r, 2].Value;
            if VarIsEmpty(pv) or VarIsNull(pv) then
              Continue;
            v := ParseMoney(VarToStr(pv));
            it := FindItem(code);
            if it <> nil then
            begin
              if it.Price <> v then
                if FData.SavePrice(code, v) then
                begin
                  LogPriceChange(code, it.Price, v, 'اکسل');
                  Inc(upd);
                end;
            end
            else
            begin
              if FData.InsertPrice(code, v) then
              begin
                LogPriceChange(code, -1, v, 'اکسل (جدید)');
                Inc(ins);
              end;
            end;
          end;
        finally
          WB.Close(False);
        end;
      finally
        XL.Quit;
        VarClear(WS);
        VarClear(WB);
        VarClear(XL);
      end;
    except
      on E: Exception do
      begin
        Application.MessageBox(PChar('خطا در خواندن اکسل:' + sLineBreak +
          E.Message), 'خطا', MB_OK or MB_ICONERROR);
        Exit;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  ReloadPrices;
  SetStatus(Format('از اکسل خوانده شد: %d به‌روزرسانی، %d ردیف جدید.',
    [upd, ins]));
  Application.MessageBox(PChar(Format
    ('آپدیت پایگاه داده از اکسل انجام شد.%s%d به‌روزرسانی، %d ردیف جدید.',
    [sLineBreak, upd, ins])), 'همگام‌سازی', MB_OK or MB_ICONINFORMATION);
end;

procedure TForm1.BtnToExcelClick(Sender: TObject);
var
  XL, WB, WS: OleVariant;
  r: Integer;
  it: TPriceItem;
  bak: string;
begin
  if FPrices.Count = 0 then
  begin
    Application.MessageBox('ابتدا قیمت‌ها را بارگذاری کنید.',
      'خطا', MB_OK or MB_ICONWARNING);
    Exit;
  end;
  if not ResolveExcelPath then
    Exit;
  bak := ChangeFileExt(FExcelPath, '') + '_' +
    FormatDateTime('yymmdd_hhnnss', Now) + '.bak';
  if not CopyFile(PChar(FExcelPath), PChar(bak), True) then
  begin
    Application.MessageBox('ساخت نسخه پشتیبان از اکسل ناموفق بود.',
      'خطا', MB_OK or MB_ICONERROR);
    Exit;
  end;
  Screen.Cursor := crHourGlass;
  try
    try
      XL := CreateOleObject('Excel.Application');
      try
        XL.Visible := False;
        XL.DisplayAlerts := False;
        WB := XL.Workbooks.Open(FExcelPath);
        try
          WS := WB.Worksheets['Price'];
          WS.Cells.Clear;
          WS.Cells[1, 1] := 'size';
          WS.Cells[1, 2] := 'price';
          r := 2;
          for it in FPrices do
          begin
            WS.Cells[r, 1] := it.Code;
            WS.Cells[r, 2] := IntToStr(it.Price);
            Inc(r);
          end;
          WB.Save;
        finally
          WB.Close(False);
        end;
      finally
        XL.Quit;
        VarClear(WS);
        VarClear(WB);
        VarClear(XL);
      end;
    except
      on E: Exception do
      begin
        Application.MessageBox(PChar('خطا در ذخیره اکسل:' + sLineBreak +
          E.Message), 'خطا', MB_OK or MB_ICONERROR);
        Exit;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  SetStatus('جدول قیمت در اکسل ذخیره شد. پشتیبان: ' + bak);
  Application.MessageBox(PChar('قیمت‌ها در اکسل ذخیره شد.' + sLineBreak +
    'پشتیبان: ' + bak), 'همگام‌سازی', MB_OK or MB_ICONINFORMATION);
end;

procedure TForm1.ShowPage(Index: Integer);
var
  i: Integer;
begin
  if (Index < 0) or (Index > High(FPages)) then
    Exit;
  FCurrentPage := Index;
  for i := 0 to High(FPages) do
  begin
    FPages[i].Visible := (i = Index);
    if Assigned(FNav[i]) then
    begin
      if i = Index then
      begin
        FNav[i].Color := C_ACCENT;
        FNav[i].Font.Style := [fsBold];
      end
      else
      begin
        FNav[i].Color := C_NAVY;
        FNav[i].Font.Style := [];
      end;
    end;
  end;
  case Index of
    0:
      UpdateStatCards;
    3:
      RefreshCart;
    5:
      if FEdPeopleSearch.CanFocus then
        FEdPeopleSearch.SetFocus;
  end;
end;

procedure TForm1.NavClick(Sender: TObject);
begin
  ShowPage(TPanel(Sender).Tag);
end;

procedure TForm1.NavEnter(Sender: TObject);
var
  idx: Integer;
begin
  idx := TPanel(Sender).Tag;
  if idx <> FCurrentPage then
    TPanel(Sender).Color := Blend(C_NAVY, clWhite, 0.10);
end;

procedure TForm1.NavLeave(Sender: TObject);
var
  idx: Integer;
begin
  idx := TPanel(Sender).Tag;
  if idx <> FCurrentPage then
    TPanel(Sender).Color := C_NAVY;
end;

procedure TForm1.BtnGoEditClick(Sender: TObject);
begin
  ShowPage(2);
end;

procedure TForm1.BtnGoCalcClick(Sender: TObject);
begin
  ShowPage(3);
end;

procedure TForm1.BtnGoMatrixClick(Sender: TObject);
begin
  ShowPage(1);
end;

procedure TForm1.UpdateConnStatus;
begin
  if FData.IsConnected then
  begin
    FLblConn.Caption := '● متصل به پایگاه داده SABZ';
    FLblConn.Font.Color := RGB(120, 230, 170);
    FBtnReconnect.Visible := False;
  end
  else
  begin
    FLblConn.Caption := '● عدم اتصال به پایگاه داده';
    FLblConn.Font.Color := RGB(255, 150, 150);
    FBtnReconnect.Visible := True;
  end;
end;

procedure TForm1.UpdateStatCards;
begin
  if FData.IsConnected then
  begin
    FLblStat[0].Caption := ToPersianDigits(IntToStr(FPrices.Count));
    FLblStat[1].Caption := ToPersianDigits(IntToStr(FData.CountPeople));
    FLblStat[2].Caption := ToPersianDigits(JalaliToday);
    FLblStat[3].Caption := 'متصل';
    FLblStat[3].Font.Color := C_SUCCESS;
  end
  else
  begin
    FLblStat[0].Caption := '—';
    FLblStat[1].Caption := '—';
    FLblStat[2].Caption := ToPersianDigits(JalaliToday);
    FLblStat[3].Caption := 'قطع';
    FLblStat[3].Font.Color := C_DANGER;
  end;
  FillLogGrid;
end;

procedure TForm1.LayoutStats(Sender: TObject);
var
  p: TPanel;
  w, i: Integer;
begin
  p := Sender as TPanel;
  w := p.ClientWidth div 4;
  for i := 0 to 3 do
    if FStatCards[i] <> nil then
      FStatCards[i].Width := w - 12;
end;

procedure TForm1.LayoutCalc(Sender: TObject);
var
  p: TPanel;
begin
  p := Sender as TPanel;
  if FCalcCatCard <> nil then
    FCalcCatCard.Width := Round(p.ClientWidth * 0.46);
end;

procedure TForm1.SetStatus(const S: string);
begin
  if Assigned(FStatus) then
    FStatus.Caption := S;
end;

procedure TForm1.BtnReconnectClick(Sender: TObject);
begin
  if not ConfirmPending(True) then
    Exit;
  FAllPeopleLoaded := False;
  Screen.Cursor := crHourGlass;
  try
    if FData.Connect then
    begin
      ReloadPrices;
      SetStatus('اتصال برقرار شد و قیمت‌ها بارگذاری شدند.');
    end
    else
    begin
      UpdateConnStatus;
      Application.MessageBox(PChar('اتصال به پایگاه داده برقرار نشد:' + sLineBreak
        + FData.LastError), 'خطای اتصال', MB_OK or MB_ICONERROR);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.ReloadPrices;
begin
  Screen.Cursor := crHourGlass;
  try
    if FData.LoadPrices(FPrices) then
    begin
      RebuildIndex;
      FOriginal.Clear;
      FPctBackup.Clear;
      FCostDraft.Clear;
      BuildOriginalMap;
      UpdateConnStatus;
      FillMatrix;
      FillPriceGrid;
      FillCostGrid;
      FillCatalog;
      UpdateStatCards;
      SetStatus('تعداد ' + IntToStr(FPrices.Count) + ' ردیف قیمت بارگذاری شد.');
    end
    else
    begin
      RebuildIndex;
      UpdateConnStatus;
      SetStatus('خطا در بارگذاری قیمت‌ها: ' + FData.LastError);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.BuildOriginalMap;
var
  it: TPriceItem;
begin
  FOriginal.Clear;
  for it in FPrices do
    FOriginal.AddOrSetValue(it.Code, it.Price);
end;

function TForm1.PriceOf(const Code: string): Int64;
var
  it: TPriceItem;
begin
  it := FindItem(Code);
  if it <> nil then
    Result := it.Price
  else
    Result := -1;
end;

function MoneyCell(v: Int64): string;
begin
  if v < 0 then
    Result := '—'
  else
    Result := FormatMoney(v);
end;

function RoundPriceValue(v, AUnit: Int64; AMode: Integer): Int64;
begin
  if (AUnit <= 1) or (v = 0) then
    Exit(v);
  case AMode of
    1:
      Result := ((v + AUnit - 1) div AUnit) * AUnit;
    2:
      Result := (v div AUnit) * AUnit;
  else
    Result := ((v + (AUnit div 2)) div AUnit) * AUnit;
  end;
end;

function ParsePercentText(const S: string; out V: Double): Boolean;
var
  t: string;
  fs: TFormatSettings;
begin
  fs := TFormatSettings.Create;
  fs.DecimalSeparator := '.';
  t := ToLatinDigits(Trim(S));
  t := StringReplace(t, '٫', '.', [rfReplaceAll]);
  t := StringReplace(t, '،', '.', [rfReplaceAll]);
  t := StringReplace(t, '٪', '', [rfReplaceAll]);
  t := StringReplace(t, '%', '', [rfReplaceAll]);
  t := StringReplace(t, '/', '.', [rfReplaceAll]);
  t := Trim(t);
  Result := TryStrToFloat(t, V, fs);
end;

procedure TForm1.FillMatrix;
var
  i: Integer;
  sz: string;
  base, n, fr, f, cost: Int64;
begin
  if FGridMatrix = nil then
    Exit;
  FUpdating := True;
  try
    FGridMatrix.RowCount := 1 + Length(MatrixSizes);
    for i := 0 to High(MatrixSizes) do
    begin
      sz := MatrixSizes[i];
      base := PriceOf(sz);
      n := PriceOf(sz + 'n');
      fr := PriceOf(sz + '_');
      f := PriceOf(sz + 'f');
      FGridMatrix.Cells[0, i + 1] := sz;
      FGridMatrix.Cells[1, i + 1] := MoneyCell(base);
      FGridMatrix.Cells[2, i + 1] := MoneyCell(n);
      FGridMatrix.Cells[3, i + 1] := MoneyCell(fr);
      if (base >= 0) and (fr >= 0) then
        FGridMatrix.Cells[4, i + 1] := MoneyCell(base + fr)
      else
        FGridMatrix.Cells[4, i + 1] := '—';
      if (n >= 0) and (fr >= 0) then
        FGridMatrix.Cells[5, i + 1] := MoneyCell(n + fr)
      else
        FGridMatrix.Cells[5, i + 1] := '—';
      FGridMatrix.Cells[6, i + 1] := MoneyCell(f);
      cost := EstimatedCostOf(sz);
      if (base > 0) and (cost > 0) then
        FGridMatrix.Cells[7, i + 1] := FormatMoney(base - cost)
      else
        FGridMatrix.Cells[7, i + 1] := '—';
    end;
  finally
    FUpdating := False;
  end;
end;

function TForm1.PriceMatchesFilter(it: TPriceItem): Boolean;
var
  term, cat: string;
begin
  term := '';
  if FEdPriceSearch <> nil then
    term := Trim(FEdPriceSearch.Text);
  cat := '';
  if (FCbCat <> nil) and (FCbCat.ItemIndex > 0) then
    cat := FCbCat.Items[FCbCat.ItemIndex];
  if (cat <> '') and (PriceCategoryNames[it.Category] <> cat) then
    Exit(False);
  term := NormalizeText(term);
  if (term <> '') and (Pos(term, NormalizeText(it.Code)) = 0) and
    (Pos(term, NormalizeText(it.DisplayName)) = 0) then
    Exit(False);
  Result := True;
end;

function TForm1.IsPriceModified(it: TPriceItem): Boolean;
var
  orig: Int64;
begin
  if FOriginal.TryGetValue(it.Code, orig) then
    Result := orig <> it.Price
  else
    Result := True;
end;

procedure TForm1.FillPriceGrid;
var
  it: TPriceItem;
  r: Integer;
  orig: Int64;
begin
  if FGridPrice = nil then
    Exit;
  FUpdating := True;
  try
    FGridPrice.RowCount := 1;
    r := 1;
    for it in FPrices do
    begin
      if not PriceMatchesFilter(it) then
        Continue;
      FGridPrice.RowCount := r + 1;
      FGridPrice.Cells[0, r] := it.Code;
      FGridPrice.Cells[1, r] := it.DisplayName;
      FGridPrice.Cells[2, r] := PriceCategoryNames[it.Category];
      FGridPrice.Cells[3, r] := FormatMoney(it.Price);
      FGridPrice.Cells[4, r] := IntToStr(it.Count);
      if not IsPriceModified(it) then
        FGridPrice.Cells[5, r] := ''
      else if FOriginal.TryGetValue(it.Code, orig) and (orig > 0) then
        FGridPrice.Cells[5, r] := Format('قبلی: %s  (%s%.1f٪)',
          [FormatMoney(orig), IfThen(it.Price >= orig, '+', ''),
          (it.Price - orig) * 100.0 / orig])
      else
        FGridPrice.Cells[5, r] := 'تغییر یافته';
      FGridPrice.Objects[0, r] := it;
      Inc(r);
    end;
    if FGridPrice.RowCount > 1 then
      FGridPrice.Row := 1;
  finally
    FUpdating := False;
  end;
  UpdateChangesLabel;
end;

procedure TForm1.UpdateChangesLabel;
var
  c: Integer;
begin
  if FLblChanges = nil then
    Exit;
  c := PendingPriceCount;
  if c = 0 then
    FLblChanges.Caption := ''
  else
    FLblChanges.Caption := IntToStr(c) + ' تغییر ذخیره‌نشده';
  if FBtnSave <> nil then
  begin
    if c > 0 then
      FBtnSave.Color := C_DANGER
    else
      FBtnSave.Color := C_ACCENT;
    FBtnSave.Tag := FBtnSave.Color;
  end;
end;

procedure TForm1.FillCatalog;
var
  it: TPriceItem;
  i, r: Integer;
  term: string;
begin
  if FGridCatalog = nil then
    Exit;
  term := '';
  if FEdCatSearch <> nil then
    term := NormalizeText(FEdCatSearch.Text);
  FUpdating := True;
  try
    FGridCatalog.RowCount := 1;
    r := 1;
    for it in FPrices do
    begin
      if (term <> '') and (Pos(term, NormalizeText(it.Code)) = 0) and
        (Pos(term, NormalizeText(it.DisplayName)) = 0) then
        Continue;
      FGridCatalog.RowCount := r + 1;
      FGridCatalog.Cells[0, r] := it.DisplayName;
      FGridCatalog.Cells[1, r] := FormatMoney(it.Price);
      FGridCatalog.Objects[0, r] := it;
      Inc(r);
    end;
    if FGridCatalog.RowCount > 1 then
      FGridCatalog.Row := 1;
  finally
    FUpdating := False;
  end;
end;

procedure TForm1.EdPriceSearchChange(Sender: TObject);
begin
  FillPriceGrid;
end;

procedure TForm1.CbCatChange(Sender: TObject);
begin
  FillPriceGrid;
end;

procedure TForm1.GridPriceSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  CanSelect := True;
  FGridPrice.EditorMode := (ACol = 3) and (ARow >= 1) and
    (FGridPrice.Objects[0, ARow] <> nil);
end;

procedure TForm1.GridPriceSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: string);
var
  it: TPriceItem;
begin
  if FUpdating then
    Exit;
  if (ACol <> 3) or (ARow < 1) then
    Exit;
  it := TPriceItem(FGridPrice.Objects[0, ARow]);
  if it = nil then
    Exit;
  if ParseMoney(Value) <> it.Price then
  begin
    it.Price := ParseMoney(Value);
    UpdateChangesLabel;
    SetStatus('قیمت «' + it.DisplayName + '» به ' + FormatToman(it.Price) +
      ' تغییر یافت. برای ذخیره، دکمه «ذخیره تغییرات» را بزنید.');
  end;
end;

procedure TForm1.BtnSaveClick(Sender: TObject);
begin
  if PendingPriceCount = 0 then
    SetStatus('تغییری برای ذخیره وجود ندارد.')
  else
    DoSavePrices;
end;

function TForm1.DoSavePrices: Boolean;
var
  it: TPriceItem;
  ok, fail: Integer;
  old: Int64;
begin
  ok := 0;
  fail := 0;
  Screen.Cursor := crHourGlass;
  try
    for it in FPrices do
    begin
      if not IsPriceModified(it) then
        Continue;
      if not FOriginal.TryGetValue(it.Code, old) then
        old := -1;
      if FData.SavePrice(it.Code, it.Price) then
      begin
        LogPriceChange(it.Code, old, it.Price, 'ویرایش');
        FOriginal.AddOrSetValue(it.Code, it.Price);
        Inc(ok);
      end
      else
        Inc(fail);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  Result := fail = 0;
  if ok > 0 then
    FPctBackup.Clear;
  FillPriceGrid;
  FillLogGrid;
  if fail = 0 then
    SetStatus(IntToStr(ok) + ' تغییر با موفقیت ذخیره شد و در تاریخچه ثبت شد.')
  else
    Application.MessageBox(PChar(IntToStr(fail) + ' ردیف ذخیره نشد.' +
      sLineBreak + FData.LastError), 'خطا در ذخیره', MB_OK or MB_ICONERROR);
end;

procedure TForm1.ApplyPctClick(Sender: TObject);
const
  RoundUnits: array [0 .. 5] of Int64 = (1, 1000, 5000, 10000, 50000, 100000);
var
  pct: Double;
  AUnit: Int64;
  mode, count, skipped, idx: Integer;
  it: TPriceItem;
  newv: Int64;
begin
  if FPrices.Count = 0 then
    Exit;
  if not ParsePercentText(FEdPercent.Text, pct) then
  begin
    Application.MessageBox('درصد وارد شده معتبر نیست. مثلاً ۲۰ یا -۱۰ یا ۱۲٫۵',
      'درصد نامعتبر', MB_OK or MB_ICONWARNING);
    Exit;
  end;
  idx := FCbRound.ItemIndex;
  if (idx < 0) or (idx > 5) then
    idx := 0;
  AUnit := RoundUnits[idx];
  mode := FCbRoundMode.ItemIndex;
  if mode < 0 then
    mode := 0;

  FPctBackup.Clear;
  count := 0;
  skipped := 0;
  Screen.Cursor := crHourGlass;
  try
    for it in FPrices do
    begin
      if not PriceMatchesFilter(it) then
        Continue;
      if it.Price <= 1 then
      begin
        Inc(skipped);
        Continue;
      end;
      FPctBackup.AddOrSetValue(it.Code, it.Price);
      newv := Round(it.Price * (1 + pct / 100.0));
      it.Price := RoundPriceValue(newv, AUnit, mode);
      Inc(count);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  FillPriceGrid;
  FillMatrix;
  FillCatalog;
  UpdateChangesLabel;
  FLblPctInfo.Caption := IntToStr(count) + ' قیمت تغییر کرد';
  if skipped > 0 then
    FLblPctInfo.Caption := FLblPctInfo.Caption + '، ' + IntToStr(skipped) + ' بدون قیمت رد شد';
  SetStatus(Format('اعمال ٪%.2f روی %d قیمت انجام شد. برای ثبت نهایی «ذخیره تغییرات» را بزنید.',
    [pct, count]));
end;

procedure TForm1.UndoPctClick(Sender: TObject);
var
  it: TPriceItem;
  v: Int64;
begin
  if FPctBackup.Count = 0 then
  begin
    SetStatus('تغییر درصدی برای بازگرداندن وجود ندارد.');
    Exit;
  end;
  for it in FPrices do
    if FPctBackup.TryGetValue(it.Code, v) then
      it.Price := v;
  FPctBackup.Clear;
  FillPriceGrid;
  FillMatrix;
  FillCatalog;
  UpdateChangesLabel;
  FLblPctInfo.Caption := 'آخرین اعمال درصد بازگردانده شد.';
  SetStatus('آخرین اعمال درصد بازگردانده شد.');
end;

procedure TForm1.GridPriceDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
  it: TPriceItem;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if (ARow >= 1) and not(gdFixed in State) then
  begin
    it := TPriceItem(g.Objects[0, ARow]);
    if (it <> nil) and IsPriceModified(it) and (ACol in [3, 5]) then
    begin
      g.Canvas.Font.Color := C_DANGER;
      g.Canvas.Font.Style := [fsBold];
    end
    else if (it <> nil) and (ACol = 3) and (it.Price <= 1) then
      g.Canvas.Font.Color := C_MUTED
    else if ACol = 0 then
      g.Canvas.Font.Color := C_ACCENT;
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

procedure TForm1.GridMatrixDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if not(gdFixed in State) then
  begin
    if (ACol >= 4) and (ACol <= 5) then
      g.Canvas.Font.Style := [fsBold];
    if ACol = 0 then
    begin
      g.Canvas.Font.Color := C_ACCENT;
      g.Canvas.Font.Style := [fsBold];
    end;
    { ستون سود: سبز برای سود، قرمز برای زیان }
    if (ACol = 7) and (g.Cells[ACol, ARow] <> '—') then
    begin
      g.Canvas.Font.Style := [fsBold];
      if StartsStr('-', g.Cells[ACol, ARow]) then
        g.Canvas.Font.Color := C_DANGER
      else
        g.Canvas.Font.Color := C_SUCCESS;
    end;
    if g.Cells[ACol, ARow] = '—' then
      g.Canvas.Font.Color := C_MUTED;
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

procedure TForm1.EdCatSearchChange(Sender: TObject);
begin
  FillCatalog;
end;

procedure TForm1.GridCatDblClick(Sender: TObject);
begin
  BtnAddClick(nil);
end;

procedure TForm1.BtnAddClick(Sender: TObject);
var
  it: TPriceItem;
  line: TCartLine;
  i, sel: Integer;
begin
  if (FGridCatalog = nil) or (FGridCatalog.Row < 1) then
    Exit;
  it := TPriceItem(FGridCatalog.Objects[0, FGridCatalog.Row]);
  if it = nil then
    Exit;
  sel := -1;
  for i := 0 to FCart.Count - 1 do
    if SameText(FCart[i].Code, it.Code) then
    begin
      Inc(FCart[i].Qty);
      sel := i;
      Break;
    end;
  if sel < 0 then
  begin
    line := TCartLine.Create;
    line.Code := it.Code;
    line.Title := it.DisplayName;
    line.Qty := 1;
    line.UnitPrice := it.Price;
    line.Cost := EstimatedCostOf(it.Code);
    sel := FCart.Add(line);
  end;
  RefreshCart(sel);
  SetStatus('«' + it.DisplayName + '» به سبد سفارش افزوده شد.');
end;

procedure TForm1.BtnRemoveClick(Sender: TObject);
var
  i: Integer;
begin
  i := FGridCart.Row;
  if (i < 1) or (i > FCart.Count) then
    Exit;
  FCart.Delete(i - 1);
  RefreshCart(Min(i - 1, FCart.Count - 1));
end;

procedure TForm1.BtnClearCartClick(Sender: TObject);
begin
  if FCart.Count = 0 then
    Exit;
  if Application.MessageBox('همه اقلام سبد سفارش پاک شود؟', 'پاک کردن سبد',
    MB_YESNO or MB_ICONQUESTION) <> IDYES then
    Exit;
  FCart.Clear;
  FEdCustomer.Text := '';
  FCustomerMobile := '';
  FEdDisc.Text := '0';
  RefreshCart;
end;

procedure TForm1.RefreshCart(ASelect: Integer);
var
  i: Integer;
  line: TCartLine;
begin
  if FGridCart = nil then
    Exit;
  FUpdating := True;
  try
    FGridCart.RowCount := 1 + FCart.Count;
    for i := 0 to FCart.Count - 1 do
    begin
      line := FCart[i];
      FGridCart.Cells[0, i + 1] := line.Title;
      FGridCart.Cells[1, i + 1] := IntToStr(line.Qty);
      FGridCart.Cells[2, i + 1] := FormatMoney(line.UnitPrice);
      FGridCart.Cells[3, i + 1] := FormatMoney(line.Cost);
      FGridCart.Cells[4, i + 1] := FormatMoney(line.Total);
      FGridCart.Objects[0, i + 1] := line;
    end;
    if (ASelect >= 0) and (ASelect < FCart.Count) then
      FGridCart.Row := ASelect + 1
    else if FGridCart.RowCount > 1 then
      FGridCart.Row := 1;
  finally
    FUpdating := False;
  end;
  RecalcCart;
end;

procedure TForm1.RecalcCart;
var
  line: TCartLine;
  subtotal, totalCost, disc, payable, profit: Int64;
begin
  subtotal := 0;
  totalCost := 0;
  for line in FCart do
  begin
    subtotal := subtotal + line.Total;
    totalCost := totalCost + (Int64(line.Qty) * line.Cost);
  end;
  disc := DiscountValue(subtotal);
  payable := subtotal - disc;
  profit := payable - totalCost;
  if FLblSubtotal <> nil then
    FLblSubtotal.Caption := FormatToman(subtotal);
  if FLblCost <> nil then
    FLblCost.Caption := FormatToman(totalCost);
  if FLblProfit <> nil then
  begin
    FLblProfit.Caption := FormatToman(profit);
    if payable > 0 then
      FLblProfit.Caption := FLblProfit.Caption +
        Format('   (حاشیه سود %.0f٪)', [profit * 100.0 / payable]);
    if profit < 0 then
      FLblProfit.Font.Color := C_DANGER
    else
      FLblProfit.Font.Color := C_SUCCESS;
  end;
  if FLblPayable <> nil then
    FLblPayable.Caption := FormatToman(payable);
end;

procedure TForm1.GridCartSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  CanSelect := True;
  FGridCart.EditorMode := ((ACol = 1) or (ACol = 2) or (ACol = 3)) and
    (ARow >= 1) and (ARow <= FCart.Count);
end;

procedure TForm1.GridCartSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: string);
var
  line: TCartLine;
begin
  if FUpdating then
    Exit;
  if (ARow < 1) or (ARow > FCart.Count) then
    Exit;
  line := FCart[ARow - 1];
  if ACol = 1 then
    line.Qty := Max(1, StrToIntDef(Trim(ToLatinDigits(Value)), 1))
  else if ACol = 2 then
    line.UnitPrice := ParseMoney(Value)
  else if ACol = 3 then
    line.Cost := ParseMoney(Value)
  else
    Exit;
  FUpdating := True;
  try
    FGridCart.Cells[4, ARow] := FormatMoney(line.Total);
  finally
    FUpdating := False;
  end;
  RecalcCart;
end;

procedure TForm1.GridCartDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if not(gdFixed in State) then
  begin
    if ACol = 3 then
      g.Canvas.Font.Color := C_MUTED;
    if ACol = 4 then
    begin
      g.Canvas.Font.Color := C_ACCENT;
      g.Canvas.Font.Style := [fsBold];
    end;
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

procedure TForm1.EdDiscChange(Sender: TObject);
begin
  if not FUpdating then
    RecalcCart;
end;

function TForm1.EstimatedCostOf(const Code: string): Int64;
var
  sz: string;
  p, s, l: Int64;
begin
  sz := SizePartOf(Code);
  p := PriceOf('p' + sz);
  if p < 0 then
    p := 0;
  s := PriceOf('s' + sz);
  if s < 0 then
    s := 0;
  l := PriceOf('l' + sz);
  if l < 0 then
    l := 0;
  case ClassifyCode(Code) of
    pcArtReprint, pcArtNew:
      Result := p + l;
    pcFrame:
      Result := s + l;
    pcItalianAlbum, pcCeremony:
      Result := p + l + s;
  else
    Result := 0;
  end;
end;

function TForm1.DiscountValue(Subtotal: Int64): Int64;
var
  t: string;
  pct: Double;
begin
  { تخفیف یا مبلغ است یا درصد (با علامت ٪ یا %) }
  t := Trim(FEdDisc.Text);
  if (Pos('%', t) > 0) or (Pos('٪', t) > 0) then
  begin
    if ParsePercentText(t, pct) then
      Result := Round(Subtotal * pct / 100.0)
    else
      Result := 0;
  end
  else
    Result := ParseMoney(t);
  if Result < 0 then
    Result := 0;
  if Result > Subtotal then
    Result := Subtotal;
end;

procedure TForm1.BtnInvTodayClick(Sender: TObject);
begin
  if FEdInvDate <> nil then
    FEdInvDate.Text := ToPersianDigits(JalaliToday);
  SetStatus('تاریخ فاکتور روی امروز تنظیم شد.');
end;

procedure TForm1.EdInvDateEnter(Sender: TObject);
begin
  if FEdInvDate <> nil then
    FEdInvDate.SelectAll;
end;

function TForm1.InvoiceDateText: string;
var
  s: string;
begin
  s := '';
  if FEdInvDate <> nil then
    s := Trim(FEdInvDate.Text);
  if s = '' then
    s := JalaliToday;
  s := StringReplace(s, #$202A, '', [rfReplaceAll]);
  s := StringReplace(s, #$202C, '', [rfReplaceAll]);
  s := StringReplace(s, #$200E, '', [rfReplaceAll]);
  s := StringReplace(s, #$200F, '', [rfReplaceAll]);
  s := ToLatinDigits(s);
  s := StringReplace(s, '-', '/', [rfReplaceAll]);
  s := StringReplace(s, '.', '/', [rfReplaceAll]);
  s := StringReplace(s, ' ', '', [rfReplaceAll]);
  s := StringReplace(s, '،', '', [rfReplaceAll]);
  Result := ToPersianDigits(s);
end;

function TForm1.InvoiceText(out DocNo: string; out Revenue, Cost, Profit,
  Discount, Payable: Int64): string;
var
  line: TCartLine;
  txt: string;
  shopMode: Boolean;
begin
  DocNo := 'SABZ-' + FormatDateTime('yymmdd-hhnnss', Now);
  Revenue := 0;
  Cost := 0;
  for line in FCart do
  begin
    Revenue := Revenue + line.Total;
    Cost := Cost + Int64(line.Qty) * line.Cost;
  end;
  Discount := DiscountValue(Revenue);
  Payable := Revenue - Discount;
  Profit := Payable - Cost;
  shopMode := (FChkShop <> nil) and FChkShop.Checked and (FShopName <> '');
  if shopMode then
    txt := 'صورت‌حساب ' + FShopName
  else
    txt := 'صورت‌حساب استودیو عکاسی سبز';
  txt := txt + sLineBreak;
  if shopMode then
  begin
    txt := txt + 'صادرشده برای استودیو عکاسی سبز' + sLineBreak;
    if FShopPhone <> '' then
      txt := txt + 'تلفن چاپخانه: ' + FShopPhone + sLineBreak;
    if FShopAddress <> '' then
      txt := txt + FShopAddress + sLineBreak;
  end;
  txt := txt + 'شماره: ' + DocNo + sLineBreak;
  txt := txt + 'تاریخ: ' + InvoiceDateText + '   ساعت: ' +
    FormatDateTime('hh:nn', Now) + sLineBreak;
  if Trim(FEdCustomer.Text) <> '' then
    txt := txt + 'مشتری: ' + Trim(FEdCustomer.Text) + sLineBreak;
  if FCustomerMobile <> '' then
    txt := txt + 'موبایل: ' + FCustomerMobile + sLineBreak;
  txt := txt + '----------------------------------------' + sLineBreak;
  for line in FCart do
    txt := txt + Format('%s ×%d = %s',
      [line.Title, line.Qty, FormatToman(line.Total)]) + sLineBreak;
  txt := txt + '----------------------------------------' + sLineBreak;
  txt := txt + 'جمع فروش: ' + FormatToman(Revenue) + sLineBreak;
  if Discount > 0 then
    txt := txt + 'تخفیف: ' + FormatToman(Discount) + sLineBreak;
  txt := txt + 'قابل پرداخت: ' + FormatToman(Payable) + sLineBreak;
  txt := txt + '(' + NumberToPersianWords(Payable) + ' تومان)' + sLineBreak;
  Result := txt;
end;

function HtmlEsc(const S: string): string;
begin
  Result := StringReplace(S, '&', '&amp;', [rfReplaceAll]);
  Result := StringReplace(Result, '<', '&lt;', [rfReplaceAll]);
  Result := StringReplace(Result, '>', '&gt;', [rfReplaceAll]);
end;

function TForm1.InvoiceHTML(const DocNo: string): string;
var
  line: TCartLine;
  rows, dummy, cust, shop, title, meta2, foot2, sign1, sign2: string;
  rev, cost, prof, disc, pay: Int64;
  n: Integer;
  shopMode: Boolean;
begin
  InvoiceText(dummy, rev, cost, prof, disc, pay);
  rows := '';
  n := 0;
  for line in FCart do
  begin
    Inc(n);
    rows := rows + Format
      ('<tr><td class="c">%s</td><td>%s</td><td class="c">%s</td>' +
      '<td>%s</td><td><b>%s</b></td></tr>' + sLineBreak,
      [ToPersianDigits(IntToStr(n)), HtmlEsc(line.Title),
      ToPersianDigits(IntToStr(line.Qty)), FormatMoney(line.UnitPrice),
      FormatMoney(line.Total)]);
  end;
  cust := '';
  if Trim(FEdCustomer.Text) <> '' then
    cust := '<div><span>مشتری:</span> ' + HtmlEsc(Trim(FEdCustomer.Text)) +
      '</div>';
  if FCustomerMobile <> '' then
    cust := cust + '<div><span>موبایل:</span> ' + HtmlEsc(FCustomerMobile) +
      '</div>';
  shop := '';
  shopMode := (FChkShop <> nil) and FChkShop.Checked and (FShopName <> '');
  if shopMode then
  begin
    title := HtmlEsc(FShopName);
    meta2 := IfThen(FShopPhone <> '', '<span>تلفن:</span> ' +
      HtmlEsc(FShopPhone) + '<br>', '') +
      IfThen(FShopAddress <> '', HtmlEsc(FShopAddress) + '<br>', '');
    foot2 := 'برای استودیو عکاسی سبز صادر شد';
    sign1 := 'مهر و امضای چاپخانه';
    sign2 := 'تأیید استودیو';
  end
  else
  begin
    title := 'استودیو عکاسی سبز';
    meta2 := '';
    foot2 := 'با تشکر از اعتماد شما — استودیو عکاسی سبز';
    sign1 := 'امضای مشتری';
    sign2 := 'مهر و امضای استودیو';
  end;
  Result :=
    '<!DOCTYPE html><html lang="fa" dir="rtl"><head><meta charset="utf-8">' +
    '<title>فاکتور ' + DocNo + '</title><style>' +
    '@page{size:A5 portrait;margin:8mm;}' +
    '*{box-sizing:border-box;}' +
    'body{font-family:Tahoma,serif;font-size:11px;color:#212529;margin:0;}' +
    '.hd{display:flex;justify-content:space-between;align-items:center;' +
    'border-bottom:3px solid #1f8a70;padding-bottom:2mm;margin-bottom:2mm;}' +
    '.hd h1{font-size:17px;margin:0;color:#1b2430;}' +
    '.hd .sub{font-size:9px;color:#1f8a70;}' +
    '.shop{margin-top:1mm;font-size:10px;color:#1b2430;background:#eef7f3;' +
    'border:1px solid #cfe8df;border-radius:1mm;padding:1mm 2mm;' +
    'display:inline-block;}' +
    '.shop span{color:#1f8a70;}' +
    '.meta{font-size:10px;color:#444;text-align:left;line-height:1.6;}' +
    '.cust{display:flex;gap:6mm;font-size:10px;margin:1mm 0 2mm;}' +
    '.cust span,.meta span{color:#888;}' +
    'table{width:100%;border-collapse:collapse;}' +
    'th,td{border:1px solid #d5d9dd;padding:3px 5px;text-align:right;}' +
    'th{background:#1b2430;color:#fff;font-weight:bold;}' +
    'tbody tr:nth-child(even){background:#f6f8f9;}' +
    '.c{text-align:center;}' +
    '.tot{margin-top:3mm;width:60%;margin-right:auto;}' +
    '.tot td{border:none;padding:2px 5px;}' +
    '.big td{font-size:14px;font-weight:bold;color:#1f8a70;' +
    'border-top:2px solid #1f8a70;}' +
    '.words{margin-top:2mm;padding:2mm;background:#eef7f3;border-radius:2mm;' +
    'font-size:10px;}' +
    '.sign{display:flex;justify-content:space-between;margin-top:10mm;' +
    'font-size:10px;color:#666;}' +
    '.foot{margin-top:6mm;font-size:9px;color:#777;text-align:center;}' +
    '</style></head><body onload="window.print()">' +
    '<div class="hd"><div><h1>' + title + '</h1>' +
    '<div class="sub">صورت‌حساب فروش</div>' + shop + '</div>' +
    '<div class="meta">' + meta2 + '<span>شماره:</span> ' + DocNo +
    '<br><span>تاریخ:</span> ' + InvoiceDateText +
    ' &nbsp; <span>ساعت:</span> ' + ToPersianDigits(FormatDateTime('hh:nn', Now)) +
    '</div></div>' +
    IfThen(cust <> '', '<div class="cust">' + cust + '</div>', '') +
    '<table><thead><tr><th class="c">#</th><th>شرح</th><th class="c">تعداد</th>' +
    '<th>قیمت واحد</th><th>جمع</th></tr></thead><tbody>' + rows +
    '</tbody></table>' +
    '<table class="tot">' +
    '<tr><td>جمع فروش:</td><td>' + FormatToman(rev) + '</td></tr>' +
    IfThen(disc > 0, '<tr><td>تخفیف:</td><td>' + FormatToman(disc) +
    '</td></tr>', '') +
    '<tr class="big"><td>قابل پرداخت:</td><td>' + FormatToman(pay) +
    '</td></tr></table>' +
    '<div class="words">مبلغ به حروف: ' + NumberToPersianWords(pay) +
    ' تومان</div>' +
    '<div class="sign"><div>' + sign1 + '</div><div>' + sign2 + '</div></div>' +
    '<div class="foot">' + foot2 + '</div>' +
    '</body></html>';
end;

procedure TForm1.BtnCopyInvoiceClick(Sender: TObject);
var
  doc: string;
  rev, cost, prof, disc, pay: Int64;
begin
  if FCart.Count = 0 then
  begin
    SetStatus('سبد سفارش خالی است.');
    Exit;
  end;
  Clipboard.AsText := InvoiceText(doc, rev, cost, prof, disc, pay);
  SetStatus('صورت‌حساب کپی شد. شماره: ' + doc);
end;

procedure TForm1.BtnSaveInvoiceClick(Sender: TObject);
var
  doc, dir, fn, html, txt: string;
  rev, cost, prof, disc, pay: Int64;
  sl: TStringList;
begin
  if FCart.Count = 0 then
  begin
    Application.MessageBox('سبد سفارش خالی است.', 'فاکتور', MB_OK or MB_ICONINFORMATION);
    Exit;
  end;
  txt := InvoiceText(doc, rev, cost, prof, disc, pay);
  html := InvoiceHTML(doc);
  dir := AppDir + 'Factures';
  ForceDirectories(dir);
  fn := dir + PathDelim + doc;
  sl := TStringList.Create;
  try
    sl.Text := txt;
    sl.SaveToFile(fn + '.txt', TEncoding.UTF8);
    sl.Text := html;
    sl.SaveToFile(fn + '.html', TEncoding.UTF8);
  finally
    sl.Free;
  end;
  SetStatus('فاکتور ذخیره شد: ' + fn + '.html');
  Application.MessageBox(PChar('فاکتور ذخیره شد در پوشه:' + sLineBreak + dir),
    'ذخیره فاکتور', MB_OK or MB_ICONINFORMATION);
end;

procedure TForm1.BtnPrintInvoiceClick(Sender: TObject);
var
  doc, dir, fn: string;
  rev, cost, prof, disc, pay: Int64;
  sl: TStringList;
begin
  if FCart.Count = 0 then
  begin
    Application.MessageBox('سبد سفارش خالی است.', 'فاکتور', MB_OK or MB_ICONINFORMATION);
    Exit;
  end;
  InvoiceText(doc, rev, cost, prof, disc, pay);
  dir := AppDir + 'Factures';
  ForceDirectories(dir);
  fn := dir + PathDelim + doc + '.html';
  sl := TStringList.Create;
  try
    sl.Text := InvoiceHTML(doc);
    sl.SaveToFile(fn, TEncoding.UTF8);
  finally
    sl.Free;
  end;
  ShellExecute(0, 'open', PChar(fn), nil, nil, SW_SHOWNORMAL);
  SetStatus('فاکتور A5 برای چاپ باز شد: ' + fn);
end;

procedure TForm1.BtnOpenFacturesClick(Sender: TObject);
var
  dir: string;
begin
  dir := AppDir + 'Factures';
  ForceDirectories(dir);
  ShellExecute(0, 'explore', PChar(dir), nil, nil, SW_SHOWNORMAL);
  SetStatus('پوشه فاکتورهای ذخیره‌شده باز شد.');
end;

procedure TForm1.BtnSvcBothClick(Sender: TObject);
begin
  AddService(1);
end;

procedure TForm1.BtnSvcFrameClick(Sender: TObject);
begin
  AddService(2);
end;

procedure TForm1.BtnSvcLamClick(Sender: TObject);
begin
  AddService(3);
end;

function TForm1.ServiceSize: string;
var
  it: TPriceItem;
  s: string;
begin
  Result := '';
  if (FGridCatalog <> nil) and (FGridCatalog.Row >= 1) then
  begin
    it := TPriceItem(FGridCatalog.Objects[0, FGridCatalog.Row]);
    if it <> nil then
    begin
      s := SizePartOf(it.Code);
      if Pos('*', s) > 0 then
        Exit(s);
    end;
  end;
  s := StringReplace(ToLatinDigits(Trim(FEdCatSearch.Text)), ' ', '',
    [rfReplaceAll]);
  if Pos('*', s) > 0 then
    Exit(s);
  s := StringReplace(ToLatinDigits(Trim(InputBox('سایز سرویس',
    'سایز را وارد کنید (مثال: 10*15):', '10*15'))), ' ', '', [rfReplaceAll]);
  if Pos('*', s) = 0 then
  begin
    SetStatus('سایز نامعتبر است.');
    Exit('');
  end;
  Result := s;
end;

procedure TForm1.AddService(Kind: Integer);
var
  sz, code, title, disp: string;
  price, cost, pf, pl: Int64;
  line: TCartLine;
  found: Boolean;
begin
  sz := ServiceSize;
  if sz = '' then
    Exit;
  disp := #$202A + StringReplace(sz, '*', ' × ', [rfReplaceAll]) + #$202C;
  pf := PriceOf('s' + sz);
  pl := PriceOf('l' + sz);
  case Kind of
    1:
      begin
        code := sz + '_';
        price := PriceOf(code);
        if price < 0 then
          if (pf >= 0) and (pl >= 0) then
            price := pf + pl;
        cost := Max(0, pf) + Max(0, pl);
        title := 'شاسی و لمینت ' + disp;
      end;
    2:
      begin
        code := 's' + sz;
        price := pf;
        cost := 0;
        title := 'شاسی ' + disp;
      end;
  else
    begin
      code := 'l' + sz;
      price := pl;
      cost := 0;
      title := 'لمینت ' + disp;
    end;
  end;
  if price < 0 then
  begin
    Application.MessageBox(PChar('قیمت «' + title +
      '» در جدول قیمت‌ها پیدا نشد.' + sLineBreak +
      'ابتدا کد مربوطه را در صفحه «ویرایش قیمت» ثبت کنید.'),
      'سرویس', MB_OK or MB_ICONWARNING);
    Exit;
  end;
  found := False;
  for line in FCart do
    if SameText(line.Code, code) then
    begin
      Inc(line.Qty);
      found := True;
      Break;
    end;
  if not found then
  begin
    line := TCartLine.Create;
    line.Code := code;
    line.Title := title;
    line.Qty := 1;
    line.UnitPrice := price;
    line.Cost := cost;
    FCart.Add(line);
  end;
  RefreshCart;
  SetStatus('«' + title + '» به سبد سفارش افزوده شد.');
end;

procedure TForm1.GridCatDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if (ACol = 1) and not(gdFixed in State) then
  begin
    g.Canvas.Font.Color := C_ACCENT;
    g.Canvas.Font.Style := [fsBold];
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

{ ------------------------------ تنظیمات ---------------------------------- }

procedure TForm1.LoadSettings;
var
  ini: TMemIniFile;
  cs: string;
begin
  FExcelPath := PriceXlsPath;
  FShopName := '';
  FShopPhone := '';
  FShopAddress := '';
  if not FileExists(AppDir + SettingsFile) then
    Exit;
  try
    ini := TMemIniFile.Create(AppDir + SettingsFile, TEncoding.UTF8);
    try
      FExcelPath := ini.ReadString('Excel', 'Path', PriceXlsPath);
      FShopName := Trim(ini.ReadString('PrintShop', 'Name', ''));
      FShopPhone := Trim(ini.ReadString('PrintShop', 'Phone', ''));
      FShopAddress := Trim(ini.ReadString('PrintShop', 'Address', ''));
      cs := Trim(ini.ReadString('Database', 'ConnectionString', ''));
      if cs <> '' then
        FData.ConnStr := cs;
    finally
      ini.Free;
    end;
  except
    FExcelPath := PriceXlsPath;
  end;
end;

procedure TForm1.SaveSettings;
var
  ini: TMemIniFile;
begin
  try
    ini := TMemIniFile.Create(AppDir + SettingsFile, TEncoding.UTF8);
    try
      ini.WriteString('Excel', 'Path', FExcelPath);
      if FShopName <> '' then
        ini.WriteString('PrintShop', 'Name', FShopName);
      if FShopPhone <> '' then
        ini.WriteString('PrintShop', 'Phone', FShopPhone);
      if FShopAddress <> '' then
        ini.WriteString('PrintShop', 'Address', FShopAddress);
      ini.UpdateFile;
    finally
      ini.Free;
    end;
  except
    SetStatus('ذخیره تنظیمات ممکن نشد (دسترسی نوشتن در پوشه برنامه؟)');
  end;
end;

function TForm1.ResolveExcelPath: Boolean;
var
  dlg: TOpenDialog;
begin
  Result := FileExists(FExcelPath);
  if Result then
    Exit;
  if Application.MessageBox(PChar('فایل اکسل قیمت پیدا نشد:' + sLineBreak +
    FExcelPath + sLineBreak + sLineBreak +
    'می‌خواهید محل فایل را انتخاب کنید؟ (مسیر برای دفعات بعد ذخیره می‌شود)'),
    'فایل اکسل', MB_YESNO or MB_ICONQUESTION) <> IDYES then
    Exit;
  dlg := TOpenDialog.Create(Self);
  try
    dlg.Title := 'انتخاب فایل اکسل قیمت (Price.xls)';
    dlg.Filter := 'فایل اکسل|*.xls;*.xlsx|همه فایل‌ها|*.*';
    dlg.Options := dlg.Options + [ofFileMustExist];
    if dlg.Execute then
    begin
      FExcelPath := dlg.FileName;
      SaveSettings;
      Result := True;
    end;
  finally
    dlg.Free;
  end;
end;

{ --------------------------- ایندکس قیمت‌ها ------------------------------ }

procedure TForm1.RebuildIndex;
var
  it: TPriceItem;
  key: string;
begin
  FIndex.Clear;
  for it in FPrices do
  begin
    key := LowerCase(Trim(it.Code));
    if not FIndex.ContainsKey(key) then
      FIndex.Add(key, it);
  end;
end;

function TForm1.FindItem(const Code: string): TPriceItem;
begin
  if not FIndex.TryGetValue(LowerCase(Trim(Code)), Result) then
    Result := nil;
end;

{ ----------------------- تغییرات ذخیره‌نشده ------------------------------ }

function TForm1.PendingPriceCount: Integer;
var
  it: TPriceItem;
begin
  Result := 0;
  for it in FPrices do
    if IsPriceModified(it) then
      Inc(Result);
end;

function TForm1.ConfirmPending(IncludeCost: Boolean): Boolean;
var
  n, costN: Integer;
begin
  n := PendingPriceCount;
  costN := 0;
  if IncludeCost then
    costN := FCostDraft.Count;
  if n + costN = 0 then
    Exit(True);
  case Application.MessageBox(PChar(Format('%d تغییر ذخیره‌نشده وجود دارد.%s' +
    'قبل از ادامه ذخیره شود؟%s(«خیر» = نادیده گرفتن تغییرات)',
    [n + costN, sLineBreak, sLineBreak])), 'تغییرات ذخیره‌نشده',
    MB_YESNOCANCEL or MB_ICONQUESTION) of
    IDYES:
      begin
        Result := True;
        if n > 0 then
          Result := DoSavePrices;
        if Result and (costN > 0) then
          Result := DoSaveCosts;
      end;
    IDNO:
      Result := True;
  else
    Result := False;
  end;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := ConfirmPending(True);
end;

{ ------------------------------ میان‌برها -------------------------------- }

procedure TForm1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  ed: TEdit;
begin
  if (Shift = []) and (Key >= VK_F1) and (Key <= VK_F6) then
  begin
    ShowPage(Key - VK_F1);
    Key := 0;
    Exit;
  end;

  if Shift = [ssCtrl] then
  begin
    case Key of
      Ord('S'):
        begin
          case FCurrentPage of
            2:
              BtnSaveClick(nil);
            3:
              BtnSaveInvoiceClick(nil);
            4:
              BtnCostSaveClick(nil);
          end;
          Key := 0;
        end;
      Ord('F'):
        begin
          case FCurrentPage of
            2:
              ed := FEdPriceSearch;
            3:
              ed := FEdCatSearch;
            5:
              ed := FEdPeopleSearch;
          else
            ed := nil;
          end;
          if (ed <> nil) and ed.CanFocus then
          begin
            ed.SetFocus;
            ed.SelectAll;
          end;
          Key := 0;
        end;
      Ord('R'):
        begin
          BtnReconnectClick(nil);
          Key := 0;
        end;
      Ord('P'):
        if FCurrentPage = 3 then
        begin
          BtnPrintInvoiceClick(nil);
          Key := 0;
        end;
    end;
    Exit;
  end;

  if Shift = [] then
  begin
    if (Key = VK_RETURN) and (ActiveControl = FGridCatalog) then
    begin
      BtnAddClick(nil);
      Key := 0;
    end
    else if (Key = VK_RETURN) and (ActiveControl = FGridPeople) then
    begin
      PickSelectedPerson(nil);
      Key := 0;
    end
    else if (Key = VK_DELETE) and (ActiveControl = FGridCart) then
    begin
      BtnRemoveClick(nil);
      Key := 0;
    end;
  end;
end;

procedure TForm1.EdSearchKeyPress(Sender: TObject; var Key: Char);
begin
  if Key <> #13 then
    Exit;
  Key := #0;
  if Sender = FEdCatSearch then
    BtnAddClick(nil)
  else if Sender = FEdPeopleSearch then
  begin
    if FPeopleTimer.Enabled then
      RunPeopleSearch
    else
      PickSelectedPerson(nil);
  end;
end;

{ ------------------------- تاریخچه تغییر قیمت ---------------------------- }

procedure TForm1.LogPriceChange(const Code: string; OldV, NewV: Int64;
  const Source: string);
var
  fn, line, oldS, pctS: string;
  fs: TFileStream;
  isNew: Boolean;
  bytes: TBytes;
  inv: TFormatSettings;
begin
  fn := AppDir + PriceLogFile;
  inv := TFormatSettings.Create;
  inv.DecimalSeparator := '.';
  if OldV < 0 then
    oldS := ''
  else
    oldS := IntToStr(OldV);
  if OldV > 0 then
    pctS := FormatFloat('0.0', (NewV - OldV) * 100.0 / OldV, inv)
  else
    pctS := '';
  line := JalaliToday + ',' + FormatDateTime('hh:nn', Now) + ',' +
    StringReplace(Code, ',', ' ', [rfReplaceAll]) + ',' + oldS + ',' +
    IntToStr(NewV) + ',' + pctS + ',' + Source + sLineBreak;
  try
    isNew := not FileExists(fn);
    if isNew then
      fs := TFileStream.Create(fn, fmCreate)
    else
    begin
      fs := TFileStream.Create(fn, fmOpenReadWrite or fmShareDenyWrite);
      fs.Seek(0, soEnd);
    end;
    try
      if isNew then
      begin
        bytes := TEncoding.UTF8.GetPreamble;
        if Length(bytes) > 0 then
          fs.WriteBuffer(bytes[0], Length(bytes));
        bytes := TEncoding.UTF8.GetBytes
          ('تاریخ,ساعت,کد,قیمت قبلی,قیمت جدید,درصد تغییر,منبع' + sLineBreak);
        fs.WriteBuffer(bytes[0], Length(bytes));
      end;
      bytes := TEncoding.UTF8.GetBytes(line);
      fs.WriteBuffer(bytes[0], Length(bytes));
    finally
      fs.Free;
    end;
  except
    { خطای ثبت تاریخچه نباید ذخیره قیمت را مختل کند (مثلاً فایل در اکسل باز است) }
  end;
end;

procedure TForm1.FillLogGrid;
var
  sl: TStringList;
  i, c, r: Integer;
  parts: TArray<string>;
  fn: string;
begin
  if FGridLog = nil then
    Exit;
  fn := AppDir + PriceLogFile;
  FGridLog.RowCount := 2;
  for c := 0 to FGridLog.ColCount - 1 do
    FGridLog.Cells[c, 1] := '';
  FGridLog.Cells[2, 1] := 'هنوز تغییری ثبت نشده';
  if not FileExists(fn) then
    Exit;
  sl := TStringList.Create;
  try
    try
      sl.LoadFromFile(fn, TEncoding.UTF8);
    except
      Exit;
    end;
    r := 0;
    for i := sl.Count - 1 downto 1 do
    begin
      if r >= 300 then
        Break;
      parts := sl[i].Split([',']);
      if Length(parts) < 6 then
        Continue;
      Inc(r);
      FGridLog.RowCount := r + 1;
      FGridLog.Cells[0, r] := ToPersianDigits(parts[0]);
      FGridLog.Cells[1, r] := ToPersianDigits(parts[1]);
      FGridLog.Cells[2, r] := parts[2];
      if parts[3] = '' then
        FGridLog.Cells[3, r] := '—'
      else
        FGridLog.Cells[3, r] := FormatMoney(StrToInt64Def(parts[3], 0));
      FGridLog.Cells[4, r] := FormatMoney(StrToInt64Def(parts[4], 0));
      if parts[5] = '' then
        FGridLog.Cells[5, r] := 'جدید'
      else if StartsStr('-', parts[5]) then
        FGridLog.Cells[5, r] := parts[5] + '٪'
      else
        FGridLog.Cells[5, r] := '+' + parts[5] + '٪';
    end;
  finally
    sl.Free;
  end;
end;

procedure TForm1.GridLogDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
  v: string;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if not(gdFixed in State) then
  begin
    v := g.Cells[ACol, ARow];
    if ACol = 2 then
      g.Canvas.Font.Color := C_ACCENT
    else if (ACol = 5) and (v <> '') then
    begin
      g.Canvas.Font.Style := [fsBold];
      if v = 'جدید' then
        g.Canvas.Font.Color := C_GOLD
      else if StartsStr('-', v) then
        g.Canvas.Font.Color := C_DANGER
      else
        g.Canvas.Font.Color := C_SUCCESS;
    end
    else if ACol = 4 then
      g.Canvas.Font.Style := [fsBold];
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

procedure TForm1.GridCostDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if not(gdFixed in State) then
  begin
    if ACol = 0 then
    begin
      g.Canvas.Font.Color := C_ACCENT;
      g.Canvas.Font.Style := [fsBold];
    end
    else if (ACol >= 1) and (ACol <= 3) and (ARow >= 1) then
    begin
      if FCostDraft.ContainsKey(CostPrefixes[ACol] + g.Cells[0, ARow]) then
      begin
        g.Canvas.Font.Color := C_DANGER;
        g.Canvas.Font.Style := [fsBold];
      end
      else if g.Cells[ACol, ARow] = '—' then
        g.Canvas.Font.Color := C_MUTED;
    end;
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

{ --------------------------- مشتریان (فازی) ------------------------------ }

procedure TForm1.BuildPeoplePage;
var
  page, card, toolbar, hdr: TPanel;
begin
  page := MakePanel(Self, C_BG);
  page.Parent := Self;
  page.Align := alClient;
  page.Visible := False;
  FPages[5] := page;

  hdr := MakePanel(page, C_BG);
  hdr.Align := alTop;
  hdr.Height := 46;
  with MakeLabel(hdr, 'مشتریان — جستجوی هوشمند', C_TEXT, 13, True) do
  begin
    Align := alClient;
    AutoSize := False;
    Layout := tlCenter;
  end;

  card := NewCard(page, '');
  card.Parent := page;
  card.Align := alClient;

  toolbar := MakePanel(card, C_CARD);
  toolbar.Align := alTop;
  toolbar.Height := 52;

  FEdPeopleSearch := MakeEdit(toolbar, EdPeopleSearchChange);
  FEdPeopleSearch.Left := 8;
  FEdPeopleSearch.Top := 10;
  FEdPeopleSearch.Width := 340;
  FEdPeopleSearch.Height := 32;
  FEdPeopleSearch.TextHint := 'نام، موبایل یا کد ملی... (غلط تایپی هم پیدا می‌شود)';
  FEdPeopleSearch.OnKeyPress := EdSearchKeyPress;

  with MakeButton(toolbar, 'انتخاب برای سفارش', C_ACCENT, clWhite,
    PickSelectedPerson) do
  begin
    Left := 356;
    Top := 10;
    Width := 150;
    Height := 32;
  end;

  FLblPeopleInfo := MakeLabel(toolbar,
    'برای جستجو تایپ کنید. دوبار کلیک یا Enter = انتخاب برای سفارش', C_MUTED, 8);
  FLblPeopleInfo.Left := 518;
  FLblPeopleInfo.Top := 18;
  FLblPeopleInfo.AutoSize := True;

  FGridPeople := MakeGrid(card);
  FGridPeople.Align := alClient;
  FGridPeople.Options := FGridPeople.Options + [goRowSelect];
  SetupColumns(FGridPeople,
    ['نام', 'جنسیت', 'موبایل', 'کد ملی', 'شغل', 'تعداد سفارش', 'تطابق'],
    [240, 70, 130, 130, 160, 90, 110]);
  FGridPeople.OnDblClick := PickSelectedPerson;
  FGridPeople.OnDrawCell := GridPeopleDrawCell;
end;

procedure TForm1.EdPeopleSearchChange(Sender: TObject);
begin
  { جستجو با کمی تأخیر تا با هر کلید پایگاه داده درگیر نشود }
  FPeopleTimer.Enabled := False;
  FPeopleTimer.Enabled := True;
end;

procedure TForm1.PeopleTimerTimer(Sender: TObject);
begin
  RunPeopleSearch;
end;

procedure TForm1.RunPeopleSearch;
var
  term: string;
  p: TPerson;
  d, maxD, r: Integer;
  fuzzy: TList<TPair<Integer, TPerson>>;
  pair: TPair<Integer, TPerson>;
  seen: TDictionary<Int64, Boolean>;
begin
  FPeopleTimer.Enabled := False;
  term := Trim(FEdPeopleSearch.Text);
  FPeopleView.Clear;
  FPeopleDist.Clear;
  if term = '' then
  begin
    FillPeopleGrid;
    FLblPeopleInfo.Caption :=
      'برای جستجو تایپ کنید. دوبار کلیک یا Enter = انتخاب برای سفارش';
    Exit;
  end;
  if not FData.IsConnected then
  begin
    FillPeopleGrid;
    FLblPeopleInfo.Caption := 'اتصال به پایگاه داده برقرار نیست.';
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try
    { ۱) جستجوی مستقیم در پایگاه داده }
    if FData.SearchPeople(ToLatinDigits(term), 200, FPeopleDb) then
      for p in FPeopleDb do
      begin
        FPeopleView.Add(p);
        FPeopleDist.Add(0);
      end;

    { ۲) اگر نتیجه کم بود: جستجوی فازی روی نام‌ها (غلط تایپی، ی/ک عربی) }
    if (FPeopleView.Count < 15) and (Length(term) >= 3) and
      not IsDigitsOnly(term) then
    begin
      if not FAllPeopleLoaded then
        FAllPeopleLoaded := FData.LoadAllPeople(FAllPeople);
      maxD := Max(1, Length(NormalizeText(term)) div 3);
      seen := TDictionary<Int64, Boolean>.Create;
      fuzzy := TList<TPair<Integer, TPerson>>.Create;
      try
        for p in FPeopleView do
          seen.AddOrSetValue(p.Id, True);
        for p in FAllPeople do
        begin
          if seen.ContainsKey(p.Id) then
            Continue;
          d := FuzzyDistance(term, p.Name);
          if d <= maxD then
            fuzzy.Add(TPair<Integer, TPerson>.Create(d, p));
        end;
        fuzzy.Sort(TComparer<TPair<Integer, TPerson>>.Construct(
          function(const A, B: TPair<Integer, TPerson>): Integer
          begin
            Result := A.Key - B.Key;
          end));
        r := 0;
        for pair in fuzzy do
        begin
          if r >= 50 then
            Break;
          FPeopleView.Add(pair.Value);
          FPeopleDist.Add(pair.Key);
          Inc(r);
        end;
      finally
        fuzzy.Free;
        seen.Free;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  FillPeopleGrid;
end;

procedure TForm1.FillPeopleGrid;
var
  i, c, exact: Integer;
  p: TPerson;
begin
  FUpdating := True;
  try
    FGridPeople.RowCount := Max(2, FPeopleView.Count + 1);
    for c := 0 to FGridPeople.ColCount - 1 do
      FGridPeople.Cells[c, 1] := '';
    FGridPeople.Objects[0, 1] := nil;
    exact := 0;
    for i := 0 to FPeopleView.Count - 1 do
    begin
      p := FPeopleView[i];
      FGridPeople.Cells[0, i + 1] := p.Name;
      FGridPeople.Cells[1, i + 1] := p.SexLabel;
      FGridPeople.Cells[2, i + 1] := Trim(p.Mobile);
      FGridPeople.Cells[3, i + 1] := Trim(p.NationalCode);
      FGridPeople.Cells[4, i + 1] := p.Job;
      FGridPeople.Cells[5, i + 1] := IntToStr(p.OrderCount);
      if FPeopleDist[i] = 0 then
      begin
        FGridPeople.Cells[6, i + 1] := 'دقیق';
        Inc(exact);
      end
      else
        FGridPeople.Cells[6, i + 1] := 'مشابه (' + IntToStr(FPeopleDist[i]) + ')';
      FGridPeople.Objects[0, i + 1] := p;
    end;
    FGridPeople.Row := 1;
  finally
    FUpdating := False;
  end;
  if Trim(FEdPeopleSearch.Text) <> '' then
  begin
    if FPeopleView.Count = 0 then
      FLblPeopleInfo.Caption := 'مشتری‌ای پیدا نشد.'
    else
      FLblPeopleInfo.Caption := Format('%d نتیجه — %d دقیق، %d مشابه',
        [FPeopleView.Count, exact, FPeopleView.Count - exact]);
  end;
end;

procedure TForm1.PickSelectedPerson(Sender: TObject);
var
  p: TPerson;
begin
  if (FGridPeople = nil) or (FGridPeople.Row < 1) then
    Exit;
  p := TPerson(FGridPeople.Objects[0, FGridPeople.Row]);
  if p = nil then
    Exit;
  FEdCustomer.Text := p.Name;
  FCustomerMobile := Trim(p.Mobile);
  ShowPage(3);
  SetStatus('مشتری «' + p.Name + '» برای سفارش انتخاب شد.');
end;

procedure TForm1.GridPeopleDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  g: TStringGrid;
begin
  g := Sender as TStringGrid;
  PrepareCell(g, ARow, State);
  if not(gdFixed in State) then
  begin
    if ACol = 0 then
      g.Canvas.Font.Style := [fsBold]
    else if (ACol = 6) and (g.Cells[ACol, ARow] <> '') then
    begin
      g.Canvas.Font.Style := [fsBold];
      if g.Cells[ACol, ARow] = 'دقیق' then
        g.Canvas.Font.Color := C_SUCCESS
      else
        g.Canvas.Font.Color := C_GOLD;
    end;
  end;
  FinishCell(g, ACol, ARow, Rect);
end;

procedure TForm1.BtnGoPeopleClick(Sender: TObject);
begin
  if (Trim(FEdPeopleSearch.Text) = '') and (Trim(FEdCustomer.Text) <> '') then
    FEdPeopleSearch.Text := Trim(FEdCustomer.Text);
  ShowPage(5);
end;

procedure TForm1.EdCustomerChange(Sender: TObject);
begin
  { نام دستی تایپ شد؛ موبایل مشتری قبلی دیگر معتبر نیست }
  FCustomerMobile := '';
end;

initialization

  C_BG := RGB(244, 246, 248);
  C_CARD := RGB(255, 255, 255);
  C_NAVY := RGB(27, 36, 48);
  C_ACCENT := RGB(31, 138, 112);
  C_ACCENT_L := RGB(150, 220, 195);
  C_GOLD := RGB(233, 168, 58);
  C_TEXT := RGB(33, 37, 41);
  C_MUTED := RGB(120, 130, 140);
  C_BORDER := RGB(222, 226, 230);
  C_ALT := RGB(248, 249, 250);
  C_SUCCESS := RGB(25, 135, 84);
  C_DANGER := RGB(200, 60, 70);
  C_SEL := RGB(214, 236, 229);

end.
