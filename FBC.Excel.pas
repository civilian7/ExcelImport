unit FBC.Excel;

interface

uses
  System.SysUtils,
  System.Classes,
  System.Generics.Collections,
  System.Variants,
  System.IOUtils,
  System.Win.ComObj,

  Winapi.Windows,
  Winapi.ShellAPI,

  Data.DB;

type
  TDatasetHelper = class helper for TDataSet
  strict private
    const
      CRLF = #13#10;
  private
    function  GetBoolean(const AName: string): Boolean;
    function  GetByte(const AName: string): TBytes;
    function  GetCurrency(const AName: string): Currency;
    function  GetDateTime(const AName: string): TDateTime;
    function  GetFloat(const AName: string): Double;
    function  GetInt64(const AName: string): Int64;
    function  GetInteger(const AName: string): Integer;
    function  GetString(const AName: string): string;
    procedure SetBoolean(const AName: string; const Value: Boolean);
    procedure SetByte(const AName: string; const Value: TBytes);
    procedure SetCurrency(const AName: string; const Value: Currency);
    procedure SetDateTime(const AName: string; const Value: TDateTime);
    procedure SetFloat(const AName: string; const Value: Double);
    procedure SetInt64(const AName: string; const Value: Int64);
    procedure SetInteger(const AName: string; const Value: Integer);
    procedure SetString(const AName, Value: string);
  public
    procedure ExportCSV(const AFileName: string; const ADelimeter: string = ','; const ALaunch: Boolean = False);
    procedure ExportExcel(const AFileName: string; const ALaunch: Boolean = False);
    procedure ForEach(const AFunc: TFunc<TDataSet, Boolean>); overload;
    procedure ForEach(const AProc: TProc<TDataSet>); overload;
    function  ToArray(const AIncludeName: Boolean = True): Variant;

    property B[const AName: string]: Boolean read GetBoolean write SetBoolean;
    property C[const AName: string]: Currency read GetCurrency write SetCurrency;
    property D[const AName: string]: TDateTime read GetDateTime write SetDateTime;
    property F[const AName: string]: Double read GetFloat write SetFloat;
    property I[const AName: string]: Integer read GetInteger write SetInteger;
    property L[const AName: string]: Int64 read GetInt64 write SetInt64;
    property S[const AName: string]: string read GetString write SetString;
    property X[const AName: string]: TBytes read GetByte write SetByte;
  end;

  TRowValues = array of array of string;

  TExcel = class
  private
    FExcel: Variant;
    FWorkBook: Variant;

    function  GetDisplayAlerts: Boolean;
    function  GetSheet(Index: Integer): Variant;
    function  GetSheetName(Index: Integer): string;
    function  GetVisible: Boolean;
    function  GetWorkSheetCount: Integer;
    procedure SetDisplayAlerts(const Value: Boolean);
    procedure SetSheetName(Index: Integer; const Value: string);
    procedure SetVisible(const Value: Boolean);
  public
    constructor Create(const AVisible: Boolean = True);
    destructor Destroy; override;

    function  AddSheet(const AName: string): Integer;
    function  ColumnName(const AIndex: Integer): string;
    function  GetValues(const AIndex: Integer = 1; const AStartRow: Integer = 1): Variant;
    procedure Open(const AFileName: string = '');
    procedure Save;
    procedure SaveAs(const AFileName: string);
    procedure Write(const AIndex: Integer; const AValue: Variant); overload;
    procedure Write(const ADataSet: TDataSet; const AIndex: Integer = 1); overload;

    property DisplayAlerts: Boolean read GetDisplayAlerts write SetDisplayAlerts;
    property Sheet[Index: Integer]: Variant read GetSheet;
    property SheetName[Index: Integer]: string read GetSheetName write SetSheetName;
    property WorkBook: Variant read FWorkBook;
    property WorkSheetCount: Integer read GetWorkSheetCount;
    property Visible: Boolean read GetVisible write SetVisible;
  end;

implementation

constructor TExcel.Create(const AVisible: Boolean);
begin
  FExcel := CreateOLEObject('Excel.Application');
  FExcel.Visible := AVisible;
  FExcel.DisplayAlerts := False;
end;

destructor TExcel.Destroy;
begin
  if not VarIsNull(FExcel) then
  begin
    FExcel.Quit;
    FExcel := Unassigned;
  end;

  inherited;
end;

function TExcel.AddSheet(const AName: string): Integer;
begin
  FWorkBook.WorkSheets.Add(Null, Null, 1, -4167);
  Result := FWorkBook.WorkSheets.Count;
  FWorkBook.WorkSheets[Result].Name := AName;
end;

function TExcel.ColumnName(const AIndex: Integer): string;
var
  LDigit: Integer;
  LColumn: Integer;
begin
  Result := '';

  LColumn := AIndex;
  while (LColumn > 0) do
  begin
    LDigit := ((LColumn -1) mod 26);
    Result := Char(65 + LDigit) + Result;
    LColumn := (LColumn - LDigit) div 26;
  end;
end;

function TExcel.GetDisplayAlerts: Boolean;
begin
  Result := FExcel.DisplayAlerts;
end;

function TExcel.GetSheet(Index: Integer): Variant;
begin
  Result := FWorkBook.WorkSheets[Index];
end;

function TExcel.GetSheetName(Index: Integer): string;
begin
  Result := FWorkBook.WorkSheets[Index].Name;
end;

function TExcel.GetValues(const AIndex: Integer; const AStartRow: Integer): Variant;
var
  LWorkSheet: Variant;
  LColCount: Integer;
  LRowCount: Integer;
  i, j: Integer;
begin
  LWorkSheet := FWorkBook.WorkSheets[AIndex];
  LColCount := LWorkSheet.UsedRange.Columns.Count;
  LRowCount := LWorkSheet.UsedRange.Rows.Count;

  Result := VarArrayCreate([0, LRowCount - AStartRow, 0, LColCount-1], varVariant);
  for i := AStartRow to LRowCount do
  begin
    for j := 1 to LColCount do
    begin
      Result[i-AStartRow, j-1] := LWorkSheet.Cells[i, j];
    end;
  end;
end;

function TExcel.GetVisible: Boolean;
begin
  Result := FExcel.Visible;
end;

function TExcel.GetWorkSheetCount: Integer;
begin
  Result := FWorkBook.WorkSheets.Count;
end;

procedure TExcel.Open(const AFileName: string);
begin
  if (AFileName <> '') then
    FWorkBook := FExcel.WorkBooks.Open(AFileName)
  else
    FWorkBook := FExcel.WorkBooks.Add;
end;

procedure TExcel.Save;
begin
  FWorkBook.SaveAs;
end;

procedure TExcel.SaveAs(const AFileName: string);
begin
  FWorkBook.SaveAs(AFileName);
end;

procedure TExcel.SetDisplayAlerts(const Value: Boolean);
begin
  FExcel.DisplayAlerts := Value;
end;

procedure TExcel.SetSheetName(Index: Integer; const Value: string);
begin
  FWorkBook.WorkSheets[Index].Name := Value;
end;

procedure TExcel.SetVisible(const Value: Boolean);
begin
  FExcel.Visible := Value;
end;

procedure TExcel.Write(const AIndex: Integer; const AValue: Variant);
var
  LColCount: Integer;
  LRowCount: Integer;
  LWorkSheet: Variant;
  LRange: Variant;
begin
  LRowCount := VarArrayHighBound(AValue, 1);
  LColCount := VarArrayHighBound(AValue, 2);

  LWorkSheet := FWorkBook.WorkSheets[AIndex];
  try
    LRange := LWorkSheet.Range[LWorkSheet.Cells[1, 1], LWorkSheet.Cells[LRowCount, LColCount]];
    LRange.Value := AValue;
  finally
    LWorkSheet := Unassigned;
  end;
end;

procedure TExcel.Write(const ADataSet: TDataSet; const AIndex: Integer);
begin
  Write(AIndex, ADataSet.ToArray);
end;

{$REGION 'TDatasetHelper'}

procedure TDatasetHelper.ExportCSV(const AFileName: string; const ADelimeter: string; const ALaunch: Boolean);
var
  LBuilder: TStringBuilder;
begin
  LBuilder := TStringBuilder.Create;
  try
    ForEach(
      procedure(ADataSet: TDataSet)
      var
        i: Integer;
      begin
        for i := 0 to ADataSet.FieldCount-1 do
        begin
          case ADataSet.Fields[i].DataType of
          ftBoolean:
            LBuilder.Append(ADataSet.Fields[i].AsBoolean.ToString(True));
          ftString:
            LBuilder.Append(ADataSet.Fields[i].AsString.QuotedString('"'));
          ftDate:
            LBuilder.Append(FormatDateTime('yyyy-mm-dd', ADataSet.Fields[i].AsDateTime));
          ftDateTime:
            LBuilder.Append(FormatDateTime('yyyy-mm-dd hh:nn:ss', ADataSet.Fields[i].AsDateTime));
          ftBlob:
            LBuilder.Append('(blob)');
          else
            LBuilder.Append(ADataSet.Fields[i].AsString);
          end;

          if (i < FieldCount-1) then
            LBuilder.Append(ADelimeter)
          else
            LBuilder.Append(CRLF);
        end;
      end
    );

    TFile.WriteAllBytes(AFileName, TEncoding.UTF8.GetBytes(LBuilder.ToString));

    if ALaunch then
      ShellExecute(0, 'open', PChar(AFileName), nil, nil, SW_SHOW);
  finally
    LBuilder.Free;
  end;
end;

procedure TDatasetHelper.ExportExcel(const AFileName: string; const ALaunch: Boolean);
begin
  var LValue := ToArray;
  try

  finally
    LValue := varNull;
  end;
end;

procedure TDatasetHelper.ForEach(const AFunc: TFunc<TDataSet, Boolean>);
begin
  var LBookMark := GetBookmark;
  try
    DisableControls;
    First;
    while not eof do
    begin
      if not AFunc(Self) then
        Break;

      Next;
    end;
  finally
    GotoBookmark(LBookmark);
  end;
end;

procedure TDatasetHelper.ForEach(const AProc: TProc<TDataSet>);
begin
  var LBookMark := GetBookmark;
  try
    DisableControls;
    First;
    while not eof do
    begin
      AProc(Self);
      Next;
    end;
  finally
    GotoBookmark(LBookmark);
  end;
end;

function TDatasetHelper.GetBoolean(const AName: string): Boolean;
begin
  Result := FieldByName(AName).AsBoolean;
end;

function TDatasetHelper.GetByte(const AName: string): TBytes;
begin
  Result := TBlobField(AName).AsBytes;
end;

function TDatasetHelper.GetCurrency(const AName: string): Currency;
begin
  Result := FieldByName(AName).AsCurrency;
end;

function TDatasetHelper.GetDateTime(const AName: string): TDateTime;
begin
  Result := FieldByName(AName).AsDateTime;
end;

function TDatasetHelper.GetFloat(const AName: string): Double;
begin
  Result := FieldByName(AName).AsFloat;
end;

function TDatasetHelper.GetInt64(const AName: string): Int64;
begin
  Result := FieldByName(AName).AsLargeInt;
end;

function TDatasetHelper.GetInteger(const AName: string): Integer;
begin
  Result := FieldByName(AName).AsInteger;
end;

function TDatasetHelper.GetString(const AName: string): string;
begin
  Result := FieldByName(AName).AsString;
end;

procedure TDatasetHelper.SetBoolean(const AName: string; const Value: Boolean);
begin
  FieldByName(AName).AsBoolean := Value;
end;

procedure TDatasetHelper.SetByte(const AName: string; const Value: TBytes);
begin
  TBlobField(Aname).AsBytes := Value;
end;

procedure TDatasetHelper.SetCurrency(const AName: string; const Value: Currency);
begin
  FieldByName(AName).AsCurrency := Value;
end;

procedure TDatasetHelper.SetDateTime(const AName: string; const Value: TDateTime);
begin
  FieldByName(AName).AsDateTime := Value;
end;

procedure TDatasetHelper.SetFloat(const AName: string; const Value: Double);
begin
  FieldByName(AName).AsFloat := Value;
end;

procedure TDatasetHelper.SetInt64(const AName: string; const Value: Int64);
begin
  FieldByName(AName).AsLargeInt := Value;
end;

procedure TDatasetHelper.SetInteger(const AName: string; const Value: Integer);
begin
  FieldByName(AName).AsInteger := Value;
end;

procedure TDatasetHelper.SetString(const AName, Value: string);
begin
  FieldByName(AName).AsString := Value;
end;

function TDatasetHelper.ToArray(const AIncludeName: Boolean): Variant;
begin
  DisableControls;
  try
    var LBookMark := GetBookmark;
    var LRow := 0;

    First;

    if AIncludeName then
    begin
      for var i := 0 to FieldCount-1 do
        Result[LRow, i] := Fields[i].FieldName;

      Inc(LRow);
    end;

    while not eof do
    begin
      for var i := 0 to FieldCount-1 do
        Result[LRow, i] := Fields[i].Value;

      Next;
      Inc(LRow);
    end;
    GotoBookmark(LBookMark);

    Result := VarArrayCreate([0, LRow-1, 0, FieldCount-1], varVariant);
  finally
    EnableControls;
  end;
end;

{$ENDREGION}

end.
