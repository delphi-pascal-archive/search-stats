unit InetThread;

interface

uses SysUtils, Variants, Classes, MSHTML, ActiveX, WinInet, Forms, Controls,
Windows, Messages, Graphics, Dialogs, StdCtrls;



type
  PMemoryStream = ^TMemoryStream;

  TErrorEvent = procedure(E: string; ThreadNum: integer) of object;
  TAcceptedEvent = procedure(const EngineConst: byte) of object;

type
  TInetThread = class(TThread)
  private
    fURL: String;
    FEngineConst: byte;
    err: string;
    MemoryStream: TMemoryStream;
    fError: TErrorEvent;
    fAccepted: TAcceptedEvent;
    procedure toError;
    procedure toAccept;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspennded: Boolean; const URL: String; Stream: PMemoryStream; EngineConst: byte);
    property OnError: TErrorEvent read fError write fError;
    property OnAcceptedEvent: TAcceptedEvent read fAccepted write fAccepted;
  end;

implementation

{ TInetThread }

constructor TInetThread.Create(CreateSuspennded: Boolean; const URL: String;
  Stream: PMemoryStream; EngineConst: byte);
begin
inherited Create(CreateSuspennded);
  FreeOnTerminate := True;
  Pointer(MemoryStream):=Stream;
  fURL:=URL;
  FEngineConst:=EngineConst;
end;

procedure TInetThread.Execute;
var
  pInet, pUrl: Pointer; //�� � ��� ��������, ���� ����� ���� ���������� ��������, ���� ���)))
  Buffer: array[0..1024] of Byte; //�����...
  BytesRead: Cardinal; //���������� ����������� ����
  i: Integer;
begin  //���� ������

  pInet := InternetOpen('InetPing', INTERNET_OPEN_TYPE_DIRECT, nil, nil, 0);
  if pInet = nil then
    begin //���� ������ �� ���������
      err := 'InternetOpen_Error'; //����� ��������� �� ������ � ����������
      Synchronize(toError); //�������� �� ������. ���� �������� ����� ����� ��������� Synchronize(AMethod: TThreadMethod)
      Exit; //������ ��� ������ ������
    end;
  try
    pUrl := InternetOpenUrl(pInet, PChar(fURL), nil, 0, INTERNET_FLAG_PRAGMA_NOCACHE or INTERNET_FLAG_RELOAD, 0);
    if pUrl = nil then //���� �� ����������� �� ���
    begin
      err := 'InternetOpenUrl_Error'; //����� ��������� �� ������ � ����������
      Synchronize(toError); //�������� �� ������. ���� �������� ����� ����� ��������� Synchronize(AMethod: TThreadMethod)
      Exit; //������ ��� ������ ������
    end;
    repeat
      FillChar(Buffer, SizeOf(Buffer), 0); //��������� ����� ������

      if InternetReadFile(pUrl, @Buffer, Length(Buffer), BytesRead) then //������ ���� �� ������ � �����
        begin
          MemoryStream.Write(Buffer, BytesRead); //����� ����� � ����� ���� ���������
        end
      else
       begin //���� ��������� �� �������
          err := 'DownLoadingFile_Error'; //����� ��������� �� ������ � ����������
          Synchronize(toError); //�������� �� ������. ���� �������� ����� ����� ��������� Synchronize(AMethod: TThreadMethod)
          Exit; //������ ��� ������ ������
       end;
    until (BytesRead = 0); //��������� ���
      MemoryStream.Position := 0; //������� ������ � ����
  finally
    if pUrl <> nil then //����������� �������?
      InternetCloseHandle(pUrl); //���������
    if pInet <> nil then //����������� �������?
      InternetCloseHandle(pInet); //���������
  end;
  if MemoryStream.Size = 0 then //�� ������ ��� ��� ����� �����, �� ���-��
  begin //���� ������ ������������ ����� 0
    err := 'DownLoadedSizeZero_Error'; //����� ��������� �� ������ � ����������
    Synchronize(toError); //�������� �� ������. ���� �������� ����� ����� ��������� Synchronize(AMethod: TThreadMethod)
    Exit; //������ ��� ������ ������
  end;
  pointer(MemoryStream) := nil; //�������� �����
if Assigned(fAccepted) then //���� ���������� ��� ���������
  Synchronize(toAccept);  //�������� �� �������� ����������
end;

procedure TInetThread.toAccept;
begin
  if Assigned(fAccepted) then
    begin
      OnAcceptedEvent(FEngineConst);
    end;
end;

procedure TInetThread.toError;
begin
 if Assigned(fError) then //���� ������� ��� ���������
    OnError(err, FEngineConst); //�������� �� ������
end;

end.
