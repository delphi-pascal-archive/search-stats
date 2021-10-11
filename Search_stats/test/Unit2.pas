unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SearchStat, StdCtrls, pngimage, ExtCtrls, Wininet;

type
  TForm2 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label11: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;
    Image6: TImage;
    Image7: TImage;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    SearchStats1: TSearchStats;
    Button2: TButton;
    Memo1: TMemo;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SearchStats1AllAcceptThreads;
    procedure SearchStats1AcceptSingleThread(const EngineConst: Byte;
      Value: Int64);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure CheckBox6Click(Sender: TObject);
    procedure CheckBox7Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure SearchStats1StopThreads;
    procedure SearchStats1Activate(const ActiveThreads: Integer);
    procedure SearchStats1Error(const ErrThreadNum, AllErrThreads: Integer;
      Error: string);
    procedure SearchStats1AcceptSingleThreadEx(const Statistic: TSingleStat);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

function GetUrlInfo(const dwInfoLevel: DWORD; const FileURL: string):
string;
var
  hSession, hFile: hInternet;
  dwBuffer: Pointer;
  dwBufferLen, dwIndex: DWORD;
begin
  Result := '';
  hSession := InternetOpen('STEROID Download',
                           INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
  if Assigned(hSession) then begin
    hFile := InternetOpenURL(hSession, PChar(FileURL), nil, 0,
                             INTERNET_FLAG_RELOAD, 0);
    dwIndex:= 0;
    dwBufferLen:= 20;
    if HttpQueryInfo(hFile, dwInfoLevel, @dwBuffer, dwBufferLen, dwIndex)
      then Result := PChar(@dwBuffer);
    if Assigned(hFile) then InternetCloseHandle(hFile);
    InternetCloseHandle(hsession);
  end;
end;


procedure TForm2.Button1Click(Sender: TObject);
begin
Memo1.Lines.Clear;
SearchStats1.URL:=Edit1.Text;
  SearchStats1.Activate;
end;

procedure TForm2.Button2Click(Sender: TObject);
begin
SearchStats1.Stop;
end;

procedure TForm2.CheckBox1Click(Sender: TObject);
begin
SearchStats1.Options.Yandex:=CheckBox1.Checked;

end;

procedure TForm2.CheckBox2Click(Sender: TObject);
begin
SearchStats1.Options.Google:=CheckBox2.Checked;
end;

procedure TForm2.CheckBox3Click(Sender: TObject);
begin
SearchStats1.Options.YahooLinks:=CheckBox3.Checked;
end;

procedure TForm2.CheckBox4Click(Sender: TObject);
begin
SearchStats1.Options.Yahoo:=CheckBox4.Checked;
end;

procedure TForm2.CheckBox5Click(Sender: TObject);
begin
SearchStats1.Options.Bing:=CheckBox5.Checked;
end;

procedure TForm2.CheckBox6Click(Sender: TObject);
begin
SearchStats1.Options.Rambler:=CheckBox6.Checked;
end;

procedure TForm2.CheckBox7Click(Sender: TObject);
begin
SearchStats1.Options.AlexaRang:=CheckBox7.Checked;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
Edit1.Text:=SearchStats1.URL;
end;


procedure TForm2.SearchStats1AcceptSingleThread(const EngineConst: Byte;
  Value: Int64);
begin
label2.Caption:=label2.Caption+IntToStr(EngineConst)+', ';
if EngineConst=1 then label5.Caption:=IntToStr(Value);
if EngineConst=2 then label7.Caption:=IntToStr(Value);
if EngineConst=3 then label9.Caption:=IntToStr(Value);
if EngineConst=4 then label11.Caption:=IntToStr(Value);
if EngineConst=5 then label13.Caption:=IntToStr(Value);
if EngineConst=6 then label15.Caption:=IntToStr(Value);
if EngineConst=7 then label17.Caption:=IntToStr(Value);

end;

procedure TForm2.SearchStats1AcceptSingleThreadEx(const Statistic: TSingleStat);
begin
Memo1.Lines.Add('Поток №'+IntToStr(Statistic.EngineConst)+'закончил работу');
Memo1.Lines.Add('Получены данные от '+Statistic.Name);
Memo1.Lines.Add('Операция: '+Statistic.Operation);
Memo1.Lines.Add('Получено значение: '+IntToStr(Statistic.Value));
Memo1.Lines.Add('+-------------------------+')
end;

procedure TForm2.SearchStats1Activate(const ActiveThreads: Integer);
begin
  label3.Caption:='Запущено '+IntToStr(ActiveThreads)+' потоков'
end;

procedure TForm2.SearchStats1AllAcceptThreads;
begin
label3.Caption:='Все потоки закончили работу';
Memo1.Lines.Add('Все потоки закончили работу')
end;

procedure TForm2.SearchStats1Error(const ErrThreadNum, AllErrThreads: Integer;
  Error: string);
begin
label3.Caption:='Ошибка потока №'+intToStr(ErrThreadNum)+'. Текст ошибки:'+Error+'. Всего ошибок в потоках '+IntToStr(AllErrThreads)
end;

procedure TForm2.SearchStats1StopThreads;
begin
label3.Caption:='Работа компонента прервана'
end;

end.
