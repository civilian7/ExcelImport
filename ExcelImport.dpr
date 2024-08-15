program ExcelImport;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {Form3},
  FBC.Excel in 'FBC.Excel.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
