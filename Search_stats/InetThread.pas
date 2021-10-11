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
  pInet, pUrl: Pointer; //не € это придумал, если можеш дать корректное описание, буду рад)))
  Buffer: array[0..1024] of Byte; //буфер...
  BytesRead: Cardinal; //количество прочитанных байт
  i: Integer;
begin  //тело потока

  pInet := InternetOpen('InetPing', INTERNET_OPEN_TYPE_DIRECT, nil, nil, 0);
  if pInet = nil then
    begin //если сесси€ не открылась
      err := 'InternetOpen_Error'; //ложим сообщение об ошибке в переменную
      Synchronize(toError); //сообщаем об ошибке. весь обратный вызов через процедуру Synchronize(AMethod: TThreadMethod)
      Exit; //больше тут делать нечего
    end;
  try
    pUrl := InternetOpenUrl(pInet, PChar(fURL), nil, 0, INTERNET_FLAG_PRAGMA_NOCACHE or INTERNET_FLAG_RELOAD, 0);
    if pUrl = nil then //если не достучались до урл
    begin
      err := 'InternetOpenUrl_Error'; //ложим сообщение об ошибке в переменную
      Synchronize(toError); //сообщаем об ошибке. весь обратный вызов через процедуру Synchronize(AMethod: TThreadMethod)
      Exit; //больше тут делать нечего
    end;
    repeat
      FillChar(Buffer, SizeOf(Buffer), 0); //заполн€ем буфер нол€ми

      if InternetReadFile(pUrl, @Buffer, Length(Buffer), BytesRead) then //читаем файл по кускам в буфер
        begin
          MemoryStream.Write(Buffer, BytesRead); //пишем буфер в поток если прочитали
        end
      else
       begin //если прочитать не удалось
          err := 'DownLoadingFile_Error'; //ложим сообщение об ошибке в переменную
          Synchronize(toError); //сообщаем об ошибке. весь обратный вызов через процедуру Synchronize(AMethod: TThreadMethod)
          Exit; //больше тут делать нечего
       end;
    until (BytesRead = 0); //прочитано все
      MemoryStream.Position := 0; //позицию потока в ноль
  finally
    if pUrl <> nil then //открывалось успешно?
      InternetCloseHandle(pUrl); //закрываем
    if pInet <> nil then //открывалось успешно?
      InternetCloseHandle(pInet); //закрываем
  end;
  if MemoryStream.Size = 0 then //не уверен что это имеет смысл, но все-же
  begin //если размер загруженного файла 0
    err := 'DownLoadedSizeZero_Error'; //ложим сообщение об ошибке в переменную
    Synchronize(toError); //сообщаем об ошибке. весь обратный вызов через процедуру Synchronize(AMethod: TThreadMethod)
    Exit; //больше тут делать нечего
  end;
  pointer(MemoryStream) := nil; //обрываем св€зь
if Assigned(fAccepted) then //если ќпјссертед был определен
  Synchronize(toAccept);  //сообщаем об успешном завершении
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
 if Assigned(fError) then //если ќп≈ггог был определен
    OnError(err, FEngineConst); //сообщаем об ошибке
end;

end.
