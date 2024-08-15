unit uMain;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  System.Generics.Collections,

  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.Grids,
  Vcl.CheckLst,
  CommCtrl,

  Data.DB,
  Datasnap.DBClient,

  FBC.Excel;

type
  TForm3 = class(TForm)
    eFileName: TEdit;
    Button1: TButton;
    StringGrid1: TStringGrid;
    CheckListBox1: TCheckListBox;
    btnSelectAll: TButton;
    btnUnselectAll: TButton;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSelectAllClick(Sender: TObject);
    procedure btnUnselectAllClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  strict private
{$REGION 'Internal'}
  type
    TCellInfo = record
      Column: Integer;
      Name: string;
      FieldName: string
    end;
  const
    MAX_CELL = 50;
    CELLS: array[0..MAX_CELL-1] of TCellInfo = (
      (Column:  0; Name: '상품주문번호'; FieldName: ''),
      (Column:  1; Name: '주문번호'; FieldName: ''),
      (Column:  2; Name: '발송처리일'; FieldName: ''),
      (Column:  3; Name: '주문상태'; FieldName: ''),
      (Column:  4; Name: '배송속성'; FieldName: ''),
      (Column:  5; Name: '풀필먼트사'; FieldName: ''),
      (Column:  6; Name: '배송방법'; FieldName: ''),
      (Column:  7; Name: '택배사'; FieldName: ''),
      (Column:  8; Name: '송장번호'; FieldName: ''),
      (Column:  9; Name: '구매확정연장상태'; FieldName: ''),
      (Column: 10; Name: '판매채널'; FieldName: ''),
      (Column: 11; Name: '구매자명'; FieldName: ''),
      (Column: 12; Name: '구매자ID'; FieldName: ''),
      (Column: 13; Name: '수취인명'; FieldName: ''),
      (Column: 14; Name: '클레임상태'; FieldName: ''),
      (Column: 15; Name: '상품번호'; FieldName: ''),
      (Column: 16; Name: '상품명'; FieldName: ''),
      (Column: 17; Name: '상품종류'; FieldName: ''),
      (Column: 18; Name: '반품안심케어'; FieldName: ''),
      (Column: 19; Name: '옵션정보'; FieldName: ''),
      (Column: 20; Name: '수량'; FieldName: ''),
      (Column: 21; Name: '상품가격'; FieldName: ''),
      (Column: 22; Name: '옵션가격'; FieldName: ''),
      (Column: 23; Name: '상품별 할인액'; FieldName: ''),
      (Column: 24; Name: '판매자 부담 할인액'; FieldName: ''),
      (Column: 25; Name: '상품별 총 주문금액'; FieldName: ''),
      (Column: 26; Name: '결제일'; FieldName: ''),
      (Column: 27; Name: '배송비 완료일'; FieldName: ''),
      (Column: 28; Name: '구매확정 요청일'; FieldName: ''),
      (Column: 29; Name: '구매확정 요청자'; FieldName: ''),
      (Column: 30; Name: '문제송장 여부'; FieldName: ''),
      (Column: 31; Name: '문제송장 등록일'; FieldName: ''),
      (Column: 32; Name: '문제송장 등록사유'; FieldName: ''),
      (Column: 33; Name: '자동구매확정예정일'; FieldName: ''),
      (Column: 34; Name: '구매확정연장 설정일'; FieldName: ''),
      (Column: 35; Name: '구매확정연장 사유'; FieldName: ''),
      (Column: 36; Name: '판매자 상품코드'; FieldName: ''),
      (Column: 37; Name: '판매자 내부코드1'; FieldName: ''),
      (Column: 38; Name: '판매자 내부코드2'; FieldName: ''),
      (Column: 39; Name: '배송비 묶음코드'; FieldName: ''),
      (Column: 40; Name: '배송비 형태'; FieldName: ''),
      (Column: 41; Name: '배송비 유형'; FieldName: ''),
      (Column: 42; Name: '배송비 합계'; FieldName: ''),
      (Column: 43; Name: '제주/도서 추가배송비'; FieldName: ''),
      (Column: 44; Name: '배송비 할인액'; FieldName: ''),
      (Column: 45; Name: '수취인 연락처1'; FieldName: ''),
      (Column: 46; Name: '수취인 연락처2'; FieldName: ''),
      (Column: 47; Name: '배송지'; FieldName: ''),
      (Column: 48; Name: '구매자 연락처'; FieldName: ''),
      (Column: 49; Name: '우편번호'; FieldName: '')
    );
{$ENDREGION}
  private
    function  GetCheckedCount: Integer;
    procedure LoadFromFile(const AFileName: string);
    procedure LoadFromURL(const AURL: string);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.btnSelectAllClick(Sender: TObject);
begin
  CheckListbox1.CheckAll(cbChecked);
end;

procedure TForm3.btnUnselectAllClick(Sender: TObject);
begin
  CheckListbox1.CheckAll(cbUnchecked);
end;

procedure TForm3.Button1Click(Sender: TObject);
begin
  if string(eFileName.Text).StartsWith('http') then
  begin
    // API로 불러오기
  end
  else
  begin
    // 엑셀에서 불러오기
    LoadFromFile(eFileName.Text);
  end;
end;

procedure TForm3.Button2Click(Sender: TObject);
begin
  var LData := TClientDataSet.Create(nil);
  try
    LData.FieldDefs.Add('name', ftString, 20);
    LData.FieldDefs.Add('age', ftInteger, 0);
    LData.FieldDefs.Add('addr', ftString, 40);
    LData.CreateDataSet;

    LData.Open;

    LData.Append;
    LData.FieldByName('name').AsString := '안영제';
    LData.FieldByName('age').AsString := '58';
    LData.FieldByName('addr').AsString := '경기도 하남시';
    LData.Post;

    LData.Append;
    LData.FieldByName('name').AsString := '박현정';
    LData.FieldByName('age').AsString := '55';
    LData.FieldByName('addr').AsString := '경기도 하남시';
    LData.Post;

    LData.ExportCSV('c:\temp\2.csv', ',', True);
  finally
    LData.Free;
  end;
end;

constructor TForm3.Create(AOwner: TComponent);
begin
  inherited;

  DTM_GETMONTHCAL
end;

destructor TForm3.Destroy;
begin
  inherited;
end;

procedure TForm3.FormCreate(Sender: TObject);
var
  LIndex: Integer;
  i: Integer;
begin
  for i := 0 to MAX_CELL-1 do
  begin
    LIndex := CheckListBox1.Items.Add(CELLS[i].Name);
    CheckListBox1.Checked[LIndex] := True;
  end;
end;

function TForm3.GetCheckedCount: Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to CheckListbox1.Count-1 do
    if CheckListBox1.Checked[i] then
      Inc(Result);
end;

procedure TForm3.LoadFromFile(const AFileName: string);
var
  LValues: Variant;
  i, j: Integer;
  LColumn: Integer;
  LRow: Integer;
begin
  var LExcel := TExcel.Create(False);
  try
    LExcel.Open(AFileName);
    LValues := LExcel.GetValues(1, 3);

    StringGrid1.BeginUpdate;
    try
      StringGrid1.ColCount := GetCheckedCount;
      StringGrid1.RowCount := VarArrayHighBound(LValues, 1) + 2;

      // title
      LColumn := 0;
      for i := 0 to MAX_CELL-1 do
      begin
        if CheckListbox1.Checked[i] then
        begin
          StringGrid1.Cells[LColumn, 0] := CELLS[i].Name;
          Inc(LColumn);
        end;
      end;

      // values
      LRow := 1;
      for i := 0 to VarArrayHighBound(LValues, 1) do
      begin
        LColumn := 0;
        for j := 0 to MAX_CELL-1 do
        begin
          if CheckListbox1.Checked[j] then
          begin
            StringGrid1.Cells[LColumn, LRow] := VarToStr(LValues[i, j]);
            Inc(LColumn);
          end;
        end;
        Inc(LRow);
      end;
    finally
      StringGrid1.EndUpdate;
    end;
  finally
    LExcel.Free;
  end;
end;

procedure TForm3.LoadFromURL(const AURL: string);
begin
  // 작업중...
end;

end.
