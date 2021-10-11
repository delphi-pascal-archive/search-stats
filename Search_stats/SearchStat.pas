unit SearchStat;

interface

uses SysUtils, Variants, Classes, MSHTML, ActiveX, WinInet, Forms, Controls,
Windows, Messages, Graphics, Dialogs, StdCtrls, InetThread;

const
  Yandex = 1;
  Google = 2;
  YahooLink = 3;
  Yahoo = 4;
  Bing = 5;
  Rambler = 6;
  Alexa  = 7;

  YaLinksPattern = '*Inlinks (*';
  YaPagesPattern = '*Pages (*';

  EngineURLs : array [1..7]of string = ('http://yandex.ru/yandsearch?text=&&site=%s',
                                        'http://www.google.ru/search?q=site:%s&&hl=ru&&lr=&&num=10',
                                        'http://siteexplorer.search.yahoo.com/search?p=%s&&bwm=i',
                                        'http://siteexplorer.search.yahoo.com/search?p=%s',
                                        'http://www.bing.com/search?q=site:%s&&filt=all',
                                        'http://nova.rambler.ru/srch?query=&&and=1&&dlang=0&&mimex=0&&st_date=&&end_date=&&news=0&&limitcontext=0&&exclude=&&filter=%s',
                                        'http://www.alexa.com/siteinfo/%s');

type
  TCharsets = (tchNone, tchUtf8, tchWin1251);

type
  TCharSet = set of Char;

type
 TOptions = class(TPersistent)
  private
    FYandex     : boolean;
    FGoogle     : boolean;
    FYahoo      : boolean;
    FBing       : boolean;
    FRambler    : boolean;
    FYahooLinks : boolean;
    FAlexaRang   : boolean;
  published
    property Yandex: Boolean read FYandex write FYandex;
    property Google: Boolean read FGoogle write FGoogle;
    property Yahoo: Boolean read FYahoo write FYahoo;
    property Bing: Boolean read FBing write FBing;
    property Rambler: Boolean read FRambler write FRambler;
    property YahooLinks: Boolean read FYahooLinks write FYahooLinks;
    property AlexaRang: Boolean read FAlexaRang write FAlexaRang;
end;

type
  TSingleStat = record
    EngineConst:byte;
    Name: string;
    Operation: string;
    Value: int64;
end;

TSingleAcceptEvent = procedure(const EngineConst: byte; Value: int64) of object;
TAllAcceptEvent = procedure of object;
TStopThreadsEvent = procedure of object;
TActivateEvent = procedure (const ActiveThreads: integer)of object;
TErrorThreadEvent = procedure (const ErrThreadNum, AllErrThreads:integer; Error: string)of object;
TSingleAcceptEventEx = procedure (const Statistic: TSingleStat)of object;

type
  TSearchStats = class(TComponent)
  private
     FURL: string;
     FYandexPages  : int64;
     FGooglePages  : int64;
     FYahooLinks   : int64;
     FYahooPages   : int64;
     FBingPages    : int64;
     FRamblerPages : int64;
     FAlexaRank    : int64;
     FActive       : boolean;
     FOptions      : TOptions;
     FThreads      : byte;
     FActiveThreads: byte;
     FAcceptThreads: byte;
     FErrorThreads : byte;
     FonSingleAcceptThread : TSingleAcceptEvent;
     FonAllAcceptThreads   : TAllAcceptEvent;
     FonStopThreads        : TStopThreadsEvent;
     FonActivate           : TActivateEvent;
     FonErrorThread        : TErrorThreadEvent;
     FonSingleAcceptEx     : TSingleAcceptEventEx;
     FPageStreams  : array [1..7]of TMemoryStream;
     FInetThreads  : array [1..7]of TInetThread;
     procedure onActivateThread;
     procedure onAcceptThread(const EngineConst: byte);
     procedure onErrorThread(ErrorStr: string; ErrThread:integer);
     procedure SetURL(cURL: string);
     procedure SetThread(EngineConst: byte);
     function  PageCharset(PageHTML: string): TCharsets;
     function  DecodeStream(Stream: TMemoryStream; Charset: TCharsets): IHTMLDocument2;
     function  ifThreadsStop: boolean;
     function  YandexParser:int64;
     function  GoogleParser:int64;
     function  YahooLinksParser: int64;
     function  YahooPagesParser: int64;
     function  BingParser: int64;
     function  RamblerParser: int64;
     function  AlexaRangRUS : int64;
  public
     Constructor Create(AOwner:TComponent);override;
     destructor Destroy;override;
     procedure Deactivate(State: boolean);
     procedure Activate;
     procedure Stop;
  published
     property Options: TOptions read FOptions write FOptions;
     property URL: string read FURL write SetURL;
     property Active: boolean read FActive write Deactivate;
     property OnAcceptSingleThread:TSingleAcceptEvent read FonSingleAcceptThread write FonSingleAcceptThread;
     property OnAllAcceptThreads:TAllAcceptEvent read FonAllAcceptThreads write FonAllAcceptThreads;
     property OnStopThreads:TStopThreadsEvent read FonStopThreads write FonStopThreads;
     property OnActivate: TActivateEvent read FonActivate write FonActivate;
     property OnError: TErrorThreadEvent read FonErrorThread write FonErrorThread;
     property OnAcceptSingleThreadEx: TSingleAcceptEventEx read FonSingleAcceptEx write FonSingleAcceptEx;
end;

procedure Register;
function StripNonConforming(const S: string; const ValidChars: TCharSet): string;
Procedure IsolateText( Const S: String; Tag1, Tag2: String; list:TStrings );
function MatchStrings(source, pattern: string): Boolean;

implementation

procedure Register;
begin
  RegisterComponents('WebDelphi.ru',[TSearchStats]);
end;

{ TSearchStats }

function StripNonConforming(const S: string; const ValidChars: TCharSet): string;
var
  DestI: Integer;
  SourceI: Integer;
begin
  SetLength(Result, Length(S));
  DestI := 0;
  for SourceI := 1 to Length(S) do
    if S[SourceI] in ValidChars then
    begin
      Inc(DestI);
      Result[DestI] := S[SourceI]
    end;
  SetLength(Result, DestI)
end;


function TSearchStats.DecodeStream(Stream: TMemoryStream;
         Charset: TCharsets): IHTMLDocument2;
var Cache: TStringList;
    V: OleVariant;
begin
  if Assigned(Stream) then
     begin
       Cache:=TStringList.Create;
       Stream.Position:=0;
       Cache.LoadFromStream(Stream);
       V:=VarArrayCreate([0, 0], varVariant);
       case Charset of
         tchUtf8:    V[0] := Utf8ToAnsi(Cache.Text);
         tchWin1251: V[0] := Cache.Text;
         tchNone:begin
                   Charset:=PageCharset(Cache.Text);
                   if Charset = tchUtf8 then
                      V[0] := Utf8ToAnsi(Cache.Text)
                   else
                      V[0] := Cache.Text;
                 end
       end;
       Result := coHTMLDocument.Create as IHTMLDocument2;
       Result.Write(PSafeArray(TVarData(V).VArray));
     end
  else
    begin
      raise Exception.Create('Поток пуст!');
      Exit;
    end;
end;

function TSearchStats.PageCharset(PageHTML: string): TCharsets;
var
  V: OleVariant;
  Doc: IHTMLDocument2;
  Meta: IHTMLElementCollection;
  Element: IHTMLElement;
  i: integer;
  cont: string;
begin
  try
    Doc:=coHTMLDocument.Create as IHTMLDocument2;
    V:=VarArrayCreate([0, 0], varVariant);
    V[0]:=PageHTML;
    Doc.Write(PSafeArray(TVarData(V).VArray));
    Meta := Doc.all.tags('meta') as IHTMLElementCollection;
    cont := '';
    for i := 0 to Meta.length - 1 do
    begin
      Element := Meta.item(i, 0) as IHTMLElement;
      if pos('content-type', LowerCase(Element.outerHTML)) > 0 then
      begin
        cont := Element.getAttribute('content', 0);
        if pos('charset', cont) > 0 then
          break
        else
          cont := '';
      end;
    end;
    if length(cont) > 0 then
    begin
      if pos('utf-8', LowerCase(cont)) > 0 then
        Result := tchUtf8
      else
        Result := tchWin1251
    end
    else
      Result := tchNone;
    Doc := nil;
  except
    raise Exception.Create
      ('PageCharset - Ошибка определения кодировки страницы');
  end;
end;

function TSearchStats.RamblerParser: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Element: IHTMLElement;
    Tags: IHTMLElementCollection;
    i:integer;
    str,attr: string;
begin
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[Rambler]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[Rambler],Charset);
  Tags:=Doc.all.tags('span')as IHTMLElementCollection;
  str:='';
  for i:=0 to Tags.length-1 do
    begin
      Element:=Tags.item(i, 0) as IHTMLElement;
      try
        if (Element.className='info')then
           str:=Element.innerText;
      except

      end;
    end;
  if length(str)>0 then
    begin
      if pos(':',str)>0 then
        begin
          Delete(str, 1,pos(':',str)+1);
          Result:=StrToInt(str)
        end
      else
        Result:=-1;
    end
  else
    Result:=-1;
end;

procedure TSearchStats.Activate;
begin
FThreads:=0;
if not FActive then
  begin
    FActive:=true;
  end
else
  Exit;
if Options.FYandex     then  SetThread(Yandex);
if Options.FGoogle     then  SetThread(Google);
if Options.FYahoo      then  SetThread(Yahoo);
if Options.FBing       then  SetThread(Bing);
if Options.FRambler    then  SetThread(Rambler);
if Options.FYahooLinks then  SetThread(YahooLink);
if Options.FAlexaRang  then  SetThread(Alexa);
if FActiveThreads=0 then FActive:=false;

if Assigned(FonActivate) then
   OnActivate(FActiveThreads);
end;

function TSearchStats.AlexaRangRUS: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Element: IHTMLElement;
    Tags: IHTMLElementCollection;
    i:integer;
    str,attr: string;
begin
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[Alexa]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[Alexa],Charset);
  Tags:=Doc.all.tags('div')as IHTMLElementCollection;
  str:='';
  for i:=0 to Tags.length-1 do
    begin
      Element:=Tags.item(i, 0) as IHTMLElement;
      try
        if (Element.className='data')then
          begin
            attr:=Element.innerHTML;
            if pos('ru.png', attr)>0 then
             str:=Element.innerText;
          end;
      except

      end;
    end;
  if length(str)>0 then
    begin
      str:=StripNonConforming(str,['0'..'9']);
      Result:=StrToInt(str)
    end
  else
    Result:=-1;
end;

function TSearchStats.BingParser: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Element: IHTMLElement;
    Tags: IHTMLElementCollection;
    i:integer;
    str,attr: string;
begin
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[Bing]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[Bing],Charset);
  Tags:=Doc.all.tags('span')as IHTMLElementCollection;
  str:='';
  for i:=0 to Tags.length-1 do
    begin
      Element:=Tags.item(i, 0) as IHTMLElement;
      try
        if (Element.className='sb_count')then
           str:=Element.innerText;
      except

      end;
    end;
  if length(str)>0 then
    begin
      if pos('из',str)>0 then
        begin
          Delete(str, 1,pos('из',str)+2);
          Result:=StrToInt(StripNonConforming(str,['0'..'9']));
        end
      else
        Result:=-1;
    end
  else
    Result:=-1;
end;

constructor TSearchStats.Create(AOwner: TComponent);
begin
  inherited;
  FOptions:=TOptions.Create;
end;


procedure TSearchStats.Deactivate(State: boolean);
var i:integer;
begin
if State then
  begin
    Activate;
    Exit;
  end;
if FActiveThreads=0 then Exit;

for i:=1 to 7 do
  begin
    if Assigned(FInetThreads[i]) then
      begin
        TerminateThread(FInetThreads[i].Handle,0);
        FreeAndNil(FPageStreams[i]);
      end;
  end;
FThreads:=0;
FActiveThreads:=0;
FAcceptThreads:=0;
FErrorThreads:=0;
end;

destructor TSearchStats.Destroy;
begin
  inherited;
end;

function TSearchStats.GoogleParser: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Tags: IHTMLElementCollection;
    Element: IHTMLElement;
    i:integer;
    Str: String;
    List:TStringList;
begin
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[Google]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[Google],Charset);
  Tags:=Doc.all.tags('p')as IHTMLElementCollection;
  for i:=0 to Tags.length-1 do
    begin
      Element:=Tags.item(i, 0) as IHTMLElement;
      str:=Element.innerText;
      if (pos('Результаты',str)>0)and(pos('из приблизительно',str)>0)
         and (pos(FURL,str)>0) then
         break;
    end;
  if length(str)>0 then
    begin
      List:=TStringList.Create;
      IsolateText(Str, 'приблизительно', FURL, list);
      Result:=StrToInt(StripNonConforming(list.Text,['0'..'9']))
    end
  else
    Result:=-1;
end;

function TSearchStats.ifThreadsStop: boolean;
begin
  Result:=(FAcceptThreads=7)or(FErrorThreads=7)or(FAcceptThreads+FErrorThreads=FThreads);
end;

Procedure IsolateText( Const S: String; Tag1, Tag2: String; list:TStrings );
Var
 pScan, pEnd, pTag1, pTag2: PChar;
 foundText: String;
 searchtext: String;
Begin
  searchtext := Uppercase(S);
  Tag1:= Uppercase( Tag1 );
  Tag2:= Uppercase( Tag2 );
  pTag1:= PChar(Tag1);
  pTag2:= PChar(Tag2);
  pScan:= PChar(searchtext);
    Repeat
      pScan:= StrPos( pScan, pTag1 );
      If pScan <> Nil Then Begin
      Inc(pScan, Length( Tag1 ));
      pEnd := StrPos( pScan, pTag2 );
      If pEnd <> Nil Then Begin
      SetString( foundText,Pchar(S) + (pScan- PChar(searchtext) ),pEnd - pScan );
      list.Add( foundText );
      pScan := pEnd + Length(tag2);
        End
        Else
          pScan := Nil;
      End;
    Until pScan = Nil;
  End;



procedure TSearchStats.onAcceptThread(const EngineConst: byte);
var Single: TSingleStat;
begin
  Dec(FActiveThreads);
  Inc(FAcceptThreads);
 case EngineConst of
    1:begin
        FYandexPages:=YandexParser;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FYandexPages);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Yandex';
            Single.Operation:='Подсчёт страниц в индеке';
            Single.Value:=FYandexPages;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
    2:begin
        FGooglePages:=GoogleParser;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FGooglePages);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Google';
            Single.Operation:='Подсчёт страниц в индеке';
            Single.Value:=FGooglePages;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
    3:begin
        FYahooLinks:=YahooLinksParser;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FYahooLinks);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Yahoo';
            Single.Operation:='Подсчёт внешних ссылок на страницы сайта';
            Single.Value:=FYahooLinks;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
    4:begin
        FYahooPages:=YahooPagesParser;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FYahooPages);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Yahoo';
            Single.Operation:='Подсчёт страниц в индексе';
            Single.Value:=FYahooPages;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
    5:begin
        FBingPages:=BingParser;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FBingPages);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Bing';
            Single.Operation:='Подсчёт страниц в индексе';
            Single.Value:=FBingPages;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
    6:begin
        FRamblerPages:=RamblerParser;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FRamblerPages);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Rambler';
            Single.Operation:='Подсчёт страниц в индексе';
            Single.Value:=FRamblerPages;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
    7:begin
        FAlexaRank:=AlexaRangRUS;
        if Assigned(FonSingleAcceptThread) then
           OnAcceptSingleThread(EngineConst, FAlexaRank);
        if Assigned(FonSingleAcceptEx) then
          begin
            Single.EngineConst:=EngineConst;
            Single.Name:='Alexa';
            Single.Operation:='Текущее положение в рейтинге';
            Single.Value:=FAlexaRank;
            OnAcceptSingleThreadEx(Single);
          end;
      end;
  end;
if ifThreadsStop then
  if Assigned(FonAllAcceptThreads) then
    begin
      OnAllAcceptThreads;
      FActive:=false;
      FThreads:=0;
      FActiveThreads:=0;
      FAcceptThreads:=0;
      FErrorThreads:=0;
    end;
end;

procedure TSearchStats.onActivateThread;
begin
  Inc(FActiveThreads);
end;

procedure TSearchStats.onErrorThread(ErrorStr: string; ErrThread:integer);
begin
  Dec(FActiveThreads);
  Inc(FErrorThreads);
  if Assigned(FonErrorThread) then
    OnError(ErrThread,FErrorThreads,ErrorStr);
  if ifThreadsStop then
    if Assigned(FonAllAcceptThreads) then
      begin
        OnAllAcceptThreads;
        FActive:=false;
        FThreads:=0;
        FActiveThreads:=0;
        FAcceptThreads:=0;
        FErrorThreads:=0;
      end;
end;

procedure TSearchStats.SetThread(EngineConst: byte);
var URLstr:string;
begin
  try
    URLstr:=Format(EngineURLs[EngineConst],[FURL]);
  except
    MessageBox(0,'Ошибка в URL. Сбор статистики прерван',PChar('Ошибка потока '+IntToStr(EngineConst)),MB_OK+MB_ICONERROR);
    Exit;
  end;
  if Assigned(FPageStreams[EngineConst]) then
    FreeAndNil(FPageStreams[EngineConst]);
  FPageStreams[EngineConst]:=TMemoryStream.Create;
  FInetThreads[EngineConst]:=TInetThread.Create(true, URLstr, Pointer(FPageStreams[EngineConst]),EngineConst);
  FInetThreads[EngineConst].OnAcceptedEvent:=onAcceptThread;
  FInetThreads[EngineConst].OnError:=onErrorThread;
  FInetThreads[EngineConst].Resume;
  FInetThreads[EngineConst].FreeOnTerminate:=true;
  Inc(FActiveThreads);
  Inc(FThreads);
end;

procedure TSearchStats.SetURL(cURL: string);
var
  aURLC: TURLComponents;
  lencurl: Cardinal;
  aURL: string;
begin
  if pos('http://', cURL) = 0 then
    cURL := 'http://' + cURL;

  FillChar(aURLC, SizeOf(TURLComponents), 0);
  with aURLC do
  begin
    lpszScheme := nil;
    dwSchemeLength := INTERNET_MAX_SCHEME_LENGTH;
    lpszHostName := nil;
    dwHostNameLength := INTERNET_MAX_HOST_NAME_LENGTH;
    lpszUserName := nil;
    dwUserNameLength := INTERNET_MAX_USER_NAME_LENGTH;
    lpszPassword := nil;
    dwPasswordLength := INTERNET_MAX_PASSWORD_LENGTH;
    lpszUrlPath := nil;
    dwUrlPathLength := INTERNET_MAX_PATH_LENGTH;
    lpszExtraInfo := nil;
    dwExtraInfoLength := INTERNET_MAX_PATH_LENGTH;
    dwStructSize := SizeOf(aURLC);
  end;
  lencurl := INTERNET_MAX_URL_LENGTH;
  SetLength(aURL, lencurl);
  InternetCanonicalizeUrl(PChar(cURL), PChar(aURL), lencurl, ICU_BROWSER_MODE);
  if InternetCrackUrl(PChar(aURL), length(aURL), 0, aURLC) then
  begin
    FURL := aURLC.lpszHostName;
    Delete(FURL, pos(aURLC.lpszUrlPath, aURLC.lpszHostName), length
        (aURLC.lpszUrlPath));
  end
  else
    raise Exception.Create('SetDomain - Ошибка WinInet #' + IntToStr
        (GetLastError));
end;

procedure TSearchStats.Stop;
var
  i: Integer;
begin
  for i:=1 to 7 do
    begin
      if Assigned(FInetThreads[i]) then
         begin
           TerminateThread(FInetThreads[i].Handle,0);
           FreeAndNil(FPageStreams[i]);
          end;
    end;
FActive:=false;
if Assigned(FonStopThreads) then
  OnStopThreads;

FThreads:=0;
FActiveThreads:=0;
FAcceptThreads:=0;
FErrorThreads:=0;

end;

function TSearchStats.YahooLinksParser: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Element: IHTMLElement;
    Tags: IHTMLElementCollection;
    i:integer;
    str: string;
    StringList: TStringList;
begin
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[YahooLink]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[YahooLink],Charset);
  StringList:=TStringList.Create;
  Tags:=Doc.all.tags('span')as IHTMLElementCollection;
  for i:=0 to Tags.length-1 do
    begin
      Element:=Tags.item(i, 0) as IHTMLElement;
      str:=Element.innerText;
      if MatchStrings(str, YaLinksPattern) then
         break;
    end;
    IsolateText(Str, '(',')', StringList);
    Result:=StrToInt(StripNonConforming(StringList.Strings[0],['0'..'9']));
end;

function TSearchStats.YahooPagesParser: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Element: IHTMLElement;
    Tags: IHTMLElementCollection;
    i:integer;
    str: string;
    StringList: TStringList;
begin
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[Yahoo]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[Yahoo],Charset);
  StringList:=TStringList.Create;
  Tags:=Doc.all.tags('span')as IHTMLElementCollection;
  for i:=0 to Tags.length-1 do
    begin
      Element:=Tags.item(i, 0) as IHTMLElement;
      str:=Element.innerText;
      if MatchStrings(str, YaPagesPattern) then
         break;
    end;
    IsolateText(Str, '(',')', StringList);
    Result:=StrToInt(StripNonConforming(StringList.Strings[0],['0'..'9']));
end;

function MatchStrings(source, pattern: string): Boolean;
var
  pSource: array[0..255] of Char;
  pPattern: array[0..255] of Char;
  function MatchPattern(element, pattern: PChar): Boolean;
    function IsPatternWild(pattern: PChar): Boolean;
    var
      t: Integer;
    begin
      Result := StrScan(pattern, '*') <> nil;
      if not Result then
        Result := StrScan(pattern, '?') <> nil;
    end;
  begin
    if 0 = StrComp(pattern, '*') then
      Result := True
    else if (element^ = Chr(0)) and (pattern^ <> Chr(0)) then
      Result := False
    else if element^ = Chr(0) then
      Result := True
    else
    begin
      case pattern^ of
        '*': if MatchPattern(element, @pattern[1]) then
            Result := True
          else
            Result := MatchPattern(@element[1], pattern);
        '?': Result := MatchPattern(@element[1], @pattern[1]);
      else
        if element^ = pattern^ then
          Result := MatchPattern(@element[1], @pattern[1])
        else
          Result := False;
      end;
    end;
  end;
begin
  StrPCopy(pSource, source);
  StrPCopy(pPattern, pattern);
  Result := MatchPattern(pSource, pPattern);
end;

function TSearchStats.YandexParser: int64;
var Charset: TCharsets;
    Doc: IHTMLDocument2;
    Cache: TStringList;
    Title: string;
    k:integer;
begin
  k:=1;
  cache:=TStringList.Create;
  Cache.LoadFromStream(FPageStreams[Yandex]);
  Charset:=PageCharset(Cache.Text);
  Doc:=coHTMLDocument.Create as IHTMLDocument2;
  Doc:=DecodeStream(FPageStreams[Yandex],Charset);
  Title:=Doc.title;
  if Length(Title)=0 then Result:=-1
    else begin
           if pos('тыс',title)>0 then
             k:=1000
           else
             if pos('млн',title)>0 then
               k:=1000000
             else
               if pos('млрд',title)>0 then
                 k:=1000000000;
         end;
try
  Title:=StripNonConforming(Title, ['0'..'9']);
  Result:=StrToInt(Title)*k;
except
  Result:=-1;
end;
end;

end.
