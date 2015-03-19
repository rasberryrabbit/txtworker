unit uLuaHttpExpr;

{$mode objfpc}{$H+}
{$I ..\lib\luadefine.inc}
{$M+}

{.$define LUA_FILE_UNICODE}

interface

uses
  Classes,SysUtils, contnrs,
  Lua,lua52;

type

  { TLuaHttpExpr }

  TLuaHttpExpr=class(TLua)
    private
      FObjList:TObjectList;
      lastError:Integer;
      lastErrMsg:string;
      dumplvl:Integer;
      FUTF8File:Boolean;
      FLastRow, FLastCol : Integer;
      FHitWorkCell:Boolean;
      procedure SetLastError(i:Integer);
      procedure ResetError;
      procedure SetColHeaders;
      procedure SetRowHeaders;
      procedure StrGrid_SetValue(col, row: Integer; const str: string;
        GridUpdate: Boolean=True);
      function NumToStr(num:Double):string;
      procedure LoadCSVFile(const Filename:string; delimiter:char=',');
      procedure LoadCSVStr(const datastr:string; delimiter:char=',');
      procedure SaveCSVFile(const Filename:string; delimiter:char=',');
      function SaveCSVStr(delimiter: char=','): string;
    public
      FuseUTF8String:Boolean;
      CSVDelimit:char;
      constructor Create(AutoRegister: Boolean=True); override;
      destructor Destroy; override;

      procedure StrGrid_clear;
      function do_string(const s:string):Integer;
    published
      // error
      function GetErrorMsg(State:TLuaState):Integer;
      function GetError(State:TLuaState):Integer;
      function Abort(State:TLuaState):Integer;
      function write(State:TLuaState):Integer;
      function Clear(State:TLuaState):Integer;
      function CloseApp(State:TLuaState):Integer;
      // character code
      function UTF8Decode(State:TLuaState):Integer;
      function UTF8Encode(State:TLuaState):Integer;
      function CRC32(State:TLuaState):Integer;
      function CRC16(State:TLuaState):Integer;
      function MD5(State:TLuaState):Integer;
      function SHA1(State:TLuaState):Integer;
      function GetLongestMatchStr(State: TLuaState): Integer;
      function GetLongestMatchStrW(State: TLuaState): Integer;
      // path, filename
      function ExtractFilePath(State:TLuaState):Integer;
      function ExtractFilename(State:TLuaState):Integer;
      function ExtractFileExt(State: TLuaState): Integer;
      // http
      function URLEncode(State:TLuaState):Integer;
      function URLDecode(State:TLuaState):Integer;
      function Base64Encode(State:TLuaState):Integer;
      function Base64Decode(State:TLuaState):Integer;
      function GetHttpText(State:TLuaState):Integer;
      function GetHttpsText(State:TLuaState):Integer;
      function PostHttpText(State:TLuaState):Integer;
      // load & save txt file
      function UTF8FileMode(State:TLuaState):Integer;
      function UTF8StringMode(State:TLuaState):Integer;
      function ReadFile(State:TLuaState):Integer;
      function WriteFile(State:TLuaState):Integer;
      function MkDir(State:TLuaState):Integer;
      function RmDir(State:TLuaState):Integer;
      function ChDir(State:TLuaState): Integer;
      function ExistFile(State:TLuaState):Integer;
      // find folder
      function Find_FolderFirst(State:TLuaState): Integer;
      function Find_FolderNext(State:TLuaState): Integer;
      function Find_FolderName(State:TLuaState): Integer;
      function Find_FolderAttr(State:TLuaState): Integer;
      function Find_FolderClose(State:TLuaState): Integer;
      // file file
      function Find_FileFirst(State:TLuaState): Integer;
      function Find_FileNext(State:TLuaState): Integer;
      function Find_FileName(State:TLuaState): Integer;
      function Find_FileAttr(State:TLuaState): Integer;
      function Find_FileClose(State:TLuaState): Integer;
      // regular expression
      function RegEx_New(State:TLuaState):Integer;
      function RegEx_Delete(State:TLuaState):Integer;

      function RegEx_Match(State:TLuaState):Integer;
      function RegEx_MatchAll(State:TLuaState):Integer;
      function RegEx_MatchAllCSV(State:TLuaState):Integer;
      function RegEx_Replace(State:TLuaState):Integer;
      // hwp
      function HWP_ReadText(State:TLuaState):Integer;
      // rss
      function RSS_Read(State:TLuaState):Integer;
      // stringgrid
      function Grid_ColRow(State:TLuaState):Integer;
      function Grid_Value(State:TLuaState):Integer;
      function Grid_AutoColumn(State:TLuaState):Integer;
      function Grid_Clear(State:TLuaState):Integer;
      function Grid_Load(State:TLuaState):Integer;
      function Grid_LoadStr(State:TLuaState):Integer;
      function Grid_Save(State:TLuaState):Integer;
      function Grid_SaveStr(State:TLuaState):Integer;
      function Grid_ToTable(State:TLuaState):Integer;
      function Grid_FromTable(State:TLuaState):Integer;
      function Grid_DeleteRow(State:TLuaState):Integer;
      function Grid_DeleteCol(State:TLuaState):Integer;
      function Grid_LoadExcel(State:TLuaState):Integer;
      function Grid_SaveExcel(State:TLuaState):Integer;
      function Grid_SortCol(State:TLuaState):Integer;
      function CSVDelimiter(State: TLuaState): Integer;
      // query box
      function InputBox(State:TLuaState):Integer;
      function QueryBoxYesNo(State:TLuaState):Integer;
      function QueryBoxYesNoCancel(State:TLuaState):Integer;
      // zip
      function Zip_Add(State:TLuaState):Integer;
      function Zip_AddText(State:TLuaState):Integer;
      function Zip_Delete(State:TLuaState):Integer;
      function Zip_Freshen(State:TLuaState):Integer;
      function Zip_Extract(State:TLuaState):Integer;
      function Zip_ExtractText(State:TLuaState):Integer;
      // key
      function Check_KeyPress(State:TLuaState):Integer;
      // fomula
      function SolveFormula(State:TLuaState):Integer;
  end;

procedure GridLoadCSVFile(const Filename: string; const delimiter: char; useUTF8Str:Boolean);
procedure GridSaveCSVFile(const Filename: string; const delimiter: char; UseUTF8str:Boolean);

implementation

uses httpsend, synacode, BRRE, uhwpfile, EasyRSS, txtworker_main,
  fpspreadsheet, fpsallformats, uluaFindFile, Controls, StdCtrls, Dialogs,
  Buttons, Forms, AbZipper, AbUnzper, AbArcTyp, AbZipTyp, uSimplefmParser,
  CsvDocument, ssl_openssl, ssl_openssl_lib;

resourcestring
  rsSyntaxErrorD = 'Syntax Error %d : %s';
  rsSyntaxError = 'Syntax Error';
  rsScriptLoadEr = 'Script load Error';
  rsRuntimeError = 'Runtime Error %d : %s';
  rsRuntimeError2 = 'Runtime Error';
  rsCallErrorD = 'Call Error %d';
  rsRuntimeError3 = 'Runtime Error: %s';
  rsAbortByUser = 'Abort by user.';
  rsInvalidRegEx = 'Invalid RegEx Instance';
  rsInvalidRegexExpr = 'Invalid Regex Expression';
  rsSInvalidPara = '%s : Invalid Parameter';


procedure GridSetColHeaders(Start:PInteger);
var
  i,l:Integer;
begin
  with txtworker_main.FormLua do begin
    if Start<>nil then
      l:=Start^
      else
        l:=1;
    if l<=WorkGrid.ColCount-1 then
      for i:=l to WorkGrid.ColCount-1 do
        WorkGrid.Cells[i,0]:=IntToStr(i-1);
    if Start<>nil then
      Start^:=WorkGrid.ColCount-1;
  end;
end;

procedure GridSetRowHeaders(Start:PInteger);
var
  i,l:Integer;
begin
  with txtworker_main.FormLua do begin
    if Start<>nil then
      l:=Start^
      else
        l:=1;
    if l<=WorkGrid.RowCount-1 then
      for i:=l to WorkGrid.RowCount-1 do
        WorkGrid.Cells[0,i]:=IntToStr(i-1);
    if Start<>nil then
      Start^:=WorkGrid.RowCount-1;
  end;
end;

procedure GridSaveCSVFile(const Filename: string; const delimiter: char; UseUTF8str:Boolean);
var
  irow: Integer;
  icol: Integer;
  iFile: TFileStream;
  doc: TCSVBuilder;
begin
  doc:=TCSVBuilder.Create;
  try
    doc.Delimiter:=delimiter;
    iFile:=TFileStream.Create(Filename,fmCreate or fmOpenWrite or fmShareDenyWrite);
    try
      doc.SetOutput(iFile);
      with txtworker_main.FormLua.WorkGrid do begin
        if (RowCount>1) and (ColCount>1) then
          for irow:=1 to RowCount-1 do begin
            for icol:=1 to ColCount-1 do begin
              if not UseUTF8str then
                 doc.AppendCell(Utf8ToAnsi(Cells[icol,irow]))
                 else
                   doc.AppendCell(Cells[icol,irow]);
            end;
            if irow<RowCount-1 then
              doc.AppendRow;
          end;
      end;
    finally
      iFile.Free;
    end;
  finally
    doc.Free;
  end;
end;

function GridSaveCSVStr(const delimiter: char; UseUTF8str:Boolean):string;
var
  irow: Integer;
  icol: Integer;
  iStr: TStringStream;
  doc: TCSVBuilder;
begin
  Result:='';
  doc:=TCSVBuilder.Create;
  try
    doc.Delimiter:=delimiter;
    iStr:=TStringStream.Create('');
    try
      doc.SetOutput(iStr);
      with txtworker_main.FormLua.WorkGrid do begin
        if (RowCount>1) and (ColCount>1) then
          for irow:=1 to RowCount-1 do begin
            for icol:=1 to ColCount-1 do begin
              if not UseUTF8str then
                 doc.AppendCell(Utf8ToAnsi(Cells[icol,irow]))
                 else
                   doc.AppendCell(Cells[icol,irow]);
            end;
            if irow<RowCount-1 then
              doc.AppendRow;
          end;
      end;
      Result:=iStr.DataString;
    finally
      iStr.Free;
    end;
  finally
    doc.Free;
  end;
end;

procedure GridLoadCSVFile(const Filename: string; const delimiter: char; useUTF8Str:Boolean);
var
  irow: Integer;
  icol: Integer;
  iFile: TFileStream;
  doc: TCSVParser;
begin
  doc:=TCSVParser.Create;
  try
    FormLua.WorkGrid.BeginUpdate;
    doc.Delimiter:=delimiter;
    iFile:=TFileStream.Create(Filename,fmOpenRead or fmShareDenyWrite);
    try
      doc.SetSource(iFile);
      while doc.ParseNextCell do begin
        icol:=doc.CurrentCol+1;
        irow:=doc.CurrentRow+1;
        if icol>=FormLua.WorkGrid.ColCount then
          FormLua.WorkGrid.ColCount:=icol+1;
        if irow>=FormLua.WorkGrid.RowCount then
          FormLua.WorkGrid.RowCount:=irow+1;
        if not useUTF8Str then
            FormLua.WorkGrid.Cells[icol,irow]:=AnsiToUtf8(doc.CurrentCellText)
          else
            FormLua.WorkGrid.Cells[icol,irow]:=doc.CurrentCellText;
      end;
      GridSetColHeaders(nil);
      GridSetRowHeaders(nil);
      FormLua.WorkGrid.AutoSizeColumns;
    finally
      iFile.Free;
    end;
  finally
    doc.Free;
    FormLua.WorkGrid.EndUpdate;
  end;
end;

procedure GridLoadCSVStr(const str: string; const delimiter: char; useUTF8Str:Boolean);
var
  irow: Integer;
  icol: Integer;
  iStr: TStringStream;
  doc: TCSVParser;
begin
  doc:=TCSVParser.Create;
  try
    FormLua.WorkGrid.BeginUpdate;
    doc.Delimiter:=delimiter;
    iStr:=TStringStream.Create(str);
    try
      doc.SetSource(iStr);
      while doc.ParseNextCell do begin
        icol:=doc.CurrentCol+1;
        irow:=doc.CurrentRow+1;
        if icol>=FormLua.WorkGrid.ColCount then
          FormLua.WorkGrid.ColCount:=icol+1;
        if irow>=FormLua.WorkGrid.RowCount then
          FormLua.WorkGrid.RowCount:=irow+1;
        if not useUTF8Str then
            FormLua.WorkGrid.Cells[icol,irow]:=AnsiToUtf8(doc.CurrentCellText)
          else
            FormLua.WorkGrid.Cells[icol,irow]:=doc.CurrentCellText;
      end;
      GridSetColHeaders(nil);
      GridSetRowHeaders(nil);
      FormLua.WorkGrid.AutoSizeColumns;
    finally
      iStr.Free;
    end;
  finally
    doc.Free;
    FormLua.WorkGrid.EndUpdate;
  end;
end;


{ TLuaHttpExpr }

procedure TLuaHttpExpr.SetLastError(i: Integer);
begin
  lastError:=i;
end;

procedure TLuaHttpExpr.ResetError;
begin
  lastError:=0;
end;

procedure TLuaHttpExpr.SetColHeaders;
begin
  GridSetColHeaders(@FLastCol);
end;

procedure TLuaHttpExpr.SetRowHeaders;
begin
  GridSetRowHeaders(@FLastRow);
end;

procedure TLuaHttpExpr.StrGrid_clear;
begin
  txtworker_main.FormLua.WorkGrid.ColCount:=1;
  txtworker_main.FormLua.WorkGrid.RowCount:=1;
  FLastRow:=1;
  FLastCol:=1;
end;

procedure TLuaHttpExpr.StrGrid_SetValue(col, row: Integer;const str: string;GridUpdate:Boolean=True);
begin
  with txtworker_main.FormLua do begin
    if col>=WorkGrid.ColCount then begin
      WorkGrid.ColCount:=col+1;
    end;
    if row>=WorkGrid.RowCount then begin
      WorkGrid.RowCount:=row+1;
    end;
    WorkGrid.Cells[col, row]:=str;
    if GridUpdate then begin
      SetColHeaders;
      SetRowHeaders;
      WorkGrid.AutoSizeColumn(col);
    end;
  end;
end;

function TLuaHttpExpr.NumToStr(num: Double): string;
begin
  if frac(num)<>0 then
    Result:=format('%f',[num])
    else
      Result:=format('%.0f',[num]);
end;

procedure TLuaHttpExpr.LoadCSVFile(const Filename: string; delimiter: char);
begin
  GridLoadCSVFile(Filename,delimiter,FuseUTF8String);
  FHitWorkCell:=False;
end;

procedure TLuaHttpExpr.LoadCSVStr(const datastr: string; delimiter: char);
begin
  GridLoadCSVStr(datastr,delimiter,FuseUTF8String);
  FHitWorkCell:=False;
end;

procedure TLuaHttpExpr.SaveCSVFile(const Filename: string; delimiter: char);
begin
  GridSaveCSVFile(Filename,delimiter,FuseUTF8String);
end;

function TLuaHttpExpr.SaveCSVStr(delimiter: char): string;
begin
  Result:=GridSaveCSVStr(delimiter,FuseUTF8String);
end;

constructor TLuaHttpExpr.Create(AutoRegister: Boolean);
begin
  inherited Create(False);
  luaL_openlibs(LuaInstance);
  if AutoRegister then
    AutoRegisterFunctions(self);
  FObjList:=TObjectList.create(True);
  lastErrMsg:='';
  lastError:=0;
  dumplvl:=0;
  FUTF8File:=DefaultSystemCodePage<>CP_UTF8; // True;
  FuseUTF8String:=True;
  FLastRow:=1;
  FLastCol:=1;
  FHitWorkCell:=False;
  if DecimalSeparator=',' then
    CSVDelimit:=';'
    else
      CSVDelimit:=',';
end;

destructor TLuaHttpExpr.Destroy;
begin
  UnregisterFunctions(self);
  FObjList.Free;
  inherited Destroy;
end;

function TLuaHttpExpr.do_string(const s: string): Integer;
var
  mc:size_t;
  msg:string;
begin
  lastError:=0;
  lastErrMsg:='';
  dumplvl:=0;
  Result:=luaL_loadstring(LuaInstance,PChar(s));
  if Result<>LUA_OK then begin
    if Result=LUA_ERRSYNTAX then begin
      msg:=lua_tolstring(LuaInstance,-1,@mc);
      if msg<>'' then
        txtworker_main.FormLua.Memo1.Lines.Add(format(rsSyntaxErrorD, [mc, msg]))
        else
          txtworker_main.FormLua.Memo1.Lines.Add(rsSyntaxError);
    end else
      txtworker_main.FormLua.Memo1.Lines.Add(rsScriptLoadEr);
  end else begin
    try
      Result:=lua_pcall(LuaInstance,0,LUA_MULTRET,0);
      if Result<>LUA_OK then begin
        if Result=LUA_ERRRUN then begin
          msg:=lua_tolstring(LuaInstance,-1,@mc);
          if msg<>'' then
            txtworker_main.FormLua.Memo1.Lines.Add(format(rsRuntimeError, [mc, msg]))
            else
              txtworker_main.FormLua.Memo1.Lines.Add(rsRuntimeError2);
        end else
          txtworker_main.FormLua.Memo1.Lines.Add(format(rsCallErrorD, [Result]));
      end;
      // auto size work cells
      if FHitWorkCell then begin
        FormLua.WorkGrid.AutoSizeColumns;
        FHitWorkCell:=False;
      end;
      if lastError<>0 then
         txtworker_main.FormLua.Memo1.Lines.Add(Format('> Error : %s',[lastErrMsg]));
    except
      on e:Exception do begin
        txtworker_main.FormLua.Memo1.Lines.Add(format(rsRuntimeError3, [e.Message]));
      end;
    end;
  end;
end;

function TLuaHttpExpr.GetErrorMsg(State: TLuaState): Integer;
begin
  Result:=1;
  lua_pop(State,lua_gettop(State));
  lua_pushstring(State,lastErrMsg);
end;

function TLuaHttpExpr.GetError(State: TLuaState): Integer;
begin
  Result:=1;
  lua_pop(State,lua_gettop(State));
  lua_pushinteger(State,lastError);
end;

function TLuaHttpExpr.Abort(State: TLuaState): Integer;
begin
  lua_pop(State,lua_gettop(State));
  raise Exception.Create(rsAbortByUser);
end;

// use instead of print
function TLuaHttpExpr.write(State: TLuaState): Integer;
var
  arg,i:Integer;
  msg,temp:string;
  nm:Double;
begin
  //ResetError;
  arg:=lua_gettop(State);
  msg:='';
  try
    for i:=1 to arg do begin
      case lua_type(State,i) of
      LUA_TBOOLEAN: if lua_toboolean(State,i) then
                   temp:='True'
                   else temp:='False';
      LUA_TNIL: temp:='nil';
      LUA_TNUMBER: begin
                     nm:=lua_tonumber(State,i);
                     temp:=NumToStr(nm);
                   end;
      LUA_TSTRING: temp:=lua_tostring(State,i);
      LUA_TTABLE: begin
                    temp:='Table';
                  end;
      LUA_TLIGHTUSERDATA,
      LUA_TUSERDATA: temp:=format('UserData : %p',[lua_touserdata(State,i)]);
      LUA_TTHREAD: temp:='Thread';
      LUA_NUMTAGS: temp:='Numtags';
      else
        temp:=lua_typename(State,i);
      end;
      if msg<>'' then
        msg:=msg+'    ';
      msg:=msg+temp;
    end;
  except
    on e:exception do begin
      lastError:=2;
      lastErrMsg:=e.Message;
    end;
  end;
  lua_pop(State,arg);
  Result:=0;
  txtworker_main.FormLua.Memo1.Lines.Add(msg);
end;

// clear messages.
function TLuaHttpExpr.Clear(State: TLuaState): Integer;
begin
  lua_pop(State,lua_gettop(State));
  txtworker_main.FormLua.Memo1.Clear;
  Result:=0;
end;

// close app
function TLuaHttpExpr.CloseApp(State: TLuaState): Integer;
begin
  lua_pop(State,lua_gettop(State));
  FormLua.Close;
  Result:=0;
end;

// param1 = str, ret = str
function TLuaHttpExpr.UTF8Decode(State: TLuaState): Integer;
var
  arg:Integer;
  src,ret:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    src:=lua_tostring(State,1)
    else
      src:='';
  lua_pop(State,arg);
  if src<>'' then
    ret:=Utf8ToAnsi(src)
    else
      ret:='';
  Result:=1;
  lua_pushstring(State,ret);
end;

// param1 = str, ret = str
function TLuaHttpExpr.UTF8Encode(State: TLuaState): Integer;
var
  arg:Integer;
  src,ret:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    src:=lua_tostring(State,1)
    else
      src:='';
  lua_pop(State,arg);
  if src<>'' then
    ret:=AnsiToUtf8(src)
    else
      ret:='';
  Result:=1;
  lua_pushstring(State,ret);
end;

// param1 = str, ret = integer
function TLuaHttpExpr.CRC32(State: TLuaState): Integer;
var
  arg,dcrc32:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  dcrc32:=synacode.Crc32(temp);
  Result:=1;
  lua_pushinteger(State,dcrc32);
end;

// param1 = str, ret = word
function TLuaHttpExpr.CRC16(State: TLuaState): Integer;
var
  arg,dcrc32:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  dcrc32:=synacode.Crc16(temp);
  Result:=1;
  lua_pushinteger(State,dcrc32);
end;

// param1 = str, ret = str
function TLuaHttpExpr.MD5(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=synacode.MD5(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

// param1 = str, ret = str
function TLuaHttpExpr.SHA1(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=synacode.SHA1(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;


function LCSsubstring(str1, str2:string; var sequence:string):integer;
var
  num: array of array of integer;
  lastsubsbegin, thissubsbegin:integer;
  i,j:integer;
  len1, len2: integer;
begin
  sequence:='';
  if (str1='') or (str2='') then
  begin
    Result:=0;
    exit;
  end;
  Result:=0;
  lastsubsbegin:=0;
  len1:=Length(str1);
  len2:=Length(str2);
  SetLength(num,len2,len1);

  for i:=1 to len1 do
  begin
    for j:=1 to len2 do
    begin
      if str1[i]<>str2[j] then
        num[j-1,i-1]:=0
        else
        begin
          if (i=1) or (j=1) then
            num[j-1,i-1]:=1
          else
            num[j-1,i-1]:=1+num[j-2,i-2];
          if num[j-1,i-1]>Result then
          begin
            Result:=num[j-1,i-1];
            thissubsbegin:=i-num[j-1,i-1]+1;
            if lastsubsbegin=thissubsbegin then
              sequence:=sequence+str1[i]
            else
            begin
              lastsubsbegin:=thissubsbegin;
              sequence:=copy(str1,lastsubsbegin,i-lastsubsbegin+1);
            end;
          end;
        end;
    end;
  end;
end;

function LCSsubstringW(str1, str2:UnicodeString; var sequence:UnicodeString):integer;
var
  num: array of array of integer;
  lastsubsbegin, thissubsbegin:integer;
  i,j:integer;
  len1, len2: integer;
begin
  sequence:='';
  if (str1='') or (str2='') then
  begin
    Result:=0;
    exit;
  end;
  Result:=0;
  lastsubsbegin:=0;
  len1:=Length(str1);
  len2:=Length(str2);
  SetLength(num,len2,len1);

  for i:=1 to len1 do
  begin
    for j:=1 to len2 do
    begin
      if str1[i]<>str2[j] then
        num[j-1,i-1]:=0
        else
        begin
          if (i=1) or (j=1) then
            num[j-1,i-1]:=1
          else
            num[j-1,i-1]:=1+num[j-2,i-2];
          if num[j-1,i-1]>Result then
          begin
            Result:=num[j-1,i-1];
            thissubsbegin:=i-num[j-1,i-1]+1;
            if lastsubsbegin=thissubsbegin then
              sequence:=sequence+str1[i]
            else
            begin
              lastsubsbegin:=thissubsbegin;
              sequence:=copy(str1,lastsubsbegin,i-lastsubsbegin+1);
            end;
          end;
        end;
    end;
  end;
end;

// param1 = str, param2 = str, ret = str
function TLuaHttpExpr.GetLongestMatchStr(State: TLuaState): Integer;
var
  arg,len:Integer;
  str1,str2,rstr:string;
begin
  arg:=lua_gettop(State);
  if arg>1 then begin
    str1:=lua_tostring(State,1);
    str2:=lua_tostring(State,2);
  end;
  lua_pop(State,arg);
  if arg=2 then begin
    len:=LCSsubstring(str1,str2,rstr);
    if len=0 then
      rstr:='';
    end else
      lastError:=1;
  Result:=1;
  lua_pushstring(State,rstr);
end;

// param1 = str, param2 = str, ret = str
function TLuaHttpExpr.GetLongestMatchStrW(State: TLuaState): Integer;
var
  arg,len:Integer;
  str1,str2,rstr:UnicodeString;
begin
  arg:=lua_gettop(State);
  if arg>1 then begin
    str1:=system.UTF8Decode(lua_tostring(State,1));
    str2:=system.UTF8Decode(lua_tostring(State,2));
  end;
  lua_pop(State,arg);
  if arg=2 then begin
    len:=LCSsubstringW(str1,str2,rstr);
    if len=0 then
      rstr:='';
    end else
      lastError:=1;
  Result:=1;
  lua_pushstring(State,system.UTF8Encode(rstr));
end;

function TLuaHttpExpr.ExtractFilePath(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=SysUtils.ExtractFilePath(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

function TLuaHttpExpr.ExtractFilename(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=SysUtils.ExtractFileName(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

function TLuaHttpExpr.ExtractFileExt(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=SysUtils.ExtractFileExt(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

// param1 = string, ret = string;
function TLuaHttpExpr.URLEncode(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=EncodeURL(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

// param1 = string, ret = string
function TLuaHttpExpr.URLDecode(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=DecodeURL(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

// param1 = str, ret = str
function TLuaHttpExpr.Base64Encode(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=synacode.EncodeBase64(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

// param1 = str, ret = str
function TLuaHttpExpr.Base64Decode(State: TLuaState): Integer;
var
  arg:Integer;
  temp:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    temp:=lua_tostring(State,1)
    else
      temp:='';
  lua_pop(State,arg);
  temp:=synacode.DecodeBase64(temp);
  Result:=1;
  lua_pushstring(State,temp);
end;

// return http string, return nil if fail.
function TLuaHttpExpr.GetHttpText(State: TLuaState): Integer;
var
  args:Integer;
  url:string;
  res:TStringStream;
  synhttp:THTTPSend;
begin
  ResetError;
  args:=lua_gettop(State);
  if args>0 then
    url:=lua_tostring(State,1)
    else
      url:='';
  lua_pop(State,args);
  if url<>'' then begin
    try
      res:=TStringStream.Create('');
      try
        synhttp:=THTTPSend.Create;
        try
          if Pos('https://',url)<>0 then begin
            synhttp.Sock.CreateWithSSL(TSSLOpenSSL);
            synhttp.Sock.SSLDoConnect;
          end;
          if synhttp.HTTPMethod('GET',url) then
            res.CopyFrom(synhttp.Document,0)
            else begin
              lastError:=2;
              lastErrMsg:=format('HTTP(S) Get Error %d : %s',[synhttp.ResultCode,url]);
            end;
        finally
          synhttp.Free;
        end;
        if lastError=0 then begin
          lua_pushstring(State,res.DataString);
        end;
      finally
        res.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

function TLuaHttpExpr.GetHttpsText(State: TLuaState): Integer;
var
  args:Integer;
  url:string;
  res:TStringStream;
  synhttp:THTTPSend;
begin
  ResetError;
  args:=lua_gettop(State);
  if args>0 then
    url:=lua_tostring(State,1)
    else
      url:='';
  lua_pop(State,args);
  if url<>'' then begin
    try
      res:=TStringStream.Create('');
      try
        synhttp:=THTTPSend.Create;
        try
          synhttp.Sock.CreateWithSSL(TSSLOpenSSL);
          synhttp.Sock.SSLDoConnect;
          if synhttp.HTTPMethod('GET',url) then
            res.CopyFrom(synhttp.Document,0)
            else begin
              lastError:=2;
              lastErrMsg:=format('HTTPS Get Error %d : %s',[synhttp.ResultCode,url]);
            end;
        finally
          synhttp.Free;
        end;
        if lastError=0 then begin
          lua_pushstring(State,res.DataString);
        end;
      finally
        res.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1=url, param2=data, ret=data
function TLuaHttpExpr.PostHttpText(State: TLuaState): Integer;
var
  args:Integer;
  url,data:string;
  res:TStringStream;
begin
  ResetError;
  args:=lua_gettop(State);
  if args>1 then begin
    url:=lua_tostring(State,1);
    data:=lua_tostring(State,2);
  end else
    url:='';
  lua_pop(State,args);
  if url<>'' then begin
    try
      res:=TStringStream.Create(data);
      try
        if HttpPostBinary(url,res) then begin
          lua_pushstring(State,res.DataString);
        end else begin
          lastError:=2;
        end;
      finally
        res.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// flag : integer;
function TLuaHttpExpr.UTF8FileMode(State: TLuaState): Integer;
var
  arg,fmode:Integer;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    fmode:=lua_tointeger(State,1);
  lua_pop(State,arg);
  if arg>0 then begin
    FUTF8File:=fmode<>0;
    Result:=0;
  end else begin
    lua_pushboolean(State,FUTF8File);
    Result:=1;
  end;
end;

// flag : integer;
function TLuaHttpExpr.UTF8StringMode(State: TLuaState): Integer;
var
  arg,fmode:Integer;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    fmode:=lua_tointeger(State,1);
  lua_pop(State,arg);
  if arg>0 then begin
    FuseUTF8String:=fmode<>0;
    Result:=0;
  end else begin
    lua_pushboolean(State,FuseUTF8String);
    Result:=1;
  end;
end;

// param1 = filename, param2 = size limit, ret = string
function TLuaHttpExpr.ReadFile(State: TLuaState): Integer;
var
  arg:Integer;
  filen:string;
  slimit:Integer;
  buf:TFileStream;
  strbuf:string;
begin
  ResetError;
  Result:=1;
  arg:=lua_gettop(State);
  if arg>0 then begin
    filen:=lua_tostring(State,1);
    if arg>1 then
      slimit:=lua_tointeger(State,2)
      else
        slimit:=-1;
  end else
    filen:='';
  lua_pop(State,arg);
  if filen<>'' then begin
    if FUTF8File then
      filen:=Utf8ToAnsi(filen);
    try
      buf:=TFileStream.Create(filen,fmOpenRead or fmShareCompat);
      try
        if (buf.Size<=slimit) or (slimit=-1) then begin
          slimit:=buf.Size;
          try
            SetLength(strbuf,slimit);
            slimit:=buf.Read(strbuf[1],slimit);
            if slimit>0 then begin
              if not FuseUTF8String then
                strbuf:=AnsiToUtf8(strbuf);
              lua_pushstring(State,strbuf);
            end;
            SetLength(strbuf,0);
          except
            on e:Exception do begin
              lastError:=3;
              lastErrMsg:=e.Message;
            end;
          end;
        end else
          lastError:=2;
      finally
        buf.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = filename, param2 = string, param3 = boolean, ret = boolean
function TLuaHttpExpr.WriteFile(State: TLuaState): Integer;
var
  arg,len:Integer;
  filen,txt:string;
  buf:TFileStream;
  bAppend:Boolean;
begin
  ResetError;
  Result:=1;
  arg:=lua_gettop(State);
  if arg>1 then begin
    filen:=lua_tostring(State,1);
    txt:=lua_tostring(State,2);
    if arg>2 then
      bAppend:=lua_toboolean(State,3)
      else
        bAppend:=False;
  end else
    filen:='';
  lua_pop(State,arg);
  if filen<>'' then begin
    if FUTF8File then
      filen:=Utf8ToAnsi(filen);
    try
      buf:=TFileStream.Create(filen,fmOpenWrite or fmCreate or fmShareDenyWrite);
      try
        if not FuseUTF8String then
          txt:=Utf8ToAnsi(txt);
        len:=Length(txt);
        if bAppend then
          buf.Position:=buf.Size;
        len:=buf.Write(txt[1],len);
        lua_pushboolean(State,True);
      finally
        buf.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
     lastError:=1;
  lua_pushboolean(State,lastError=0);
end;

function TLuaHttpExpr.MkDir(State: TLuaState): Integer;
var
  arg:Integer;
  fname:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
      else
        fname:='';
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    system.mkdir(fname);
  end else
    lastError:=1;
  Result:=0;
end;

function TLuaHttpExpr.RmDir(State: TLuaState): Integer;
var
  arg:Integer;
  fname:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
      else
        fname:='';
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    system.rmdir(fname);
  end else
    lastError:=1;
  Result:=0;
end;

function TLuaHttpExpr.ChDir(State:TLuaState): Integer;
var
  arg:Integer;
  fname:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
      else
        fname:='';
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    system.chdir(fname);
  end else
    lastError:=1;
  Result:=0;
end;

// param1: filename, ret : boolean
function TLuaHttpExpr.ExistFile(State: TLuaState): Integer;
var
  arg:Integer;
  fname:string;
  fret:Boolean;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    fret:=FileExists(fname);
  end else
    fret:=False;
  Result:=1;
  lua_pushboolean(State,fret);
end;

// param1=folder, param2=recursive, ret = userdata
function TLuaHttpExpr.Find_FolderFirst(State: TLuaState): Integer;
var
  arg:Integer;
  fname:string;
  recur:Boolean;
  fobj:TLuaFindFolder;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  if arg>1 then
    recur:=lua_tointeger(State,2)<>0
    else
      recur:=True;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    try
      fobj:=TLuaFindFolder.Create(fname,recur);
      try
        lua_pushlightuserdata(State,Pointer(fobj));
        FObjList.Add(fobj);
      except
        on e:Exception do begin
          fobj.Free;
          lastError:=3;
          lastErrMsg:=e.Message;
        end;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1=obj, ret = bool
function TLuaHttpExpr.Find_FolderNext(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFolder;
  bret:boolean;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFolder(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if (fobj<>nil) and (FObjList.IndexOf(fobj)<>-1) then
    bret:=fobj.Next
    else
      bret:=False;
  Result:=1;
  lua_pushboolean(State,bret);
end;

function TLuaHttpExpr.Find_FolderName(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFolder;
  sret:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFolder(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if (fobj<>nil) and (FObjList.IndexOf(fobj)<>-1) then begin
    sret:=fobj.GetFullName;
    if FUTF8File then
      sret:=AnsiToUtf8(sret);
    end else
      sret:='';
  Result:=1;
  lua_pushstring(State,sret);
end;

// param1 = obj, ret = attr
function TLuaHttpExpr.Find_FolderAttr(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFolder;
  rattr:Integer;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFolder(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if (fobj<>nil) and (FObjList.IndexOf(fobj)<>-1) then
    rattr:=fobj.GetAttr
    else
      rattr:=0;
  Result:=1;
  lua_pushinteger(State,rattr);
end;

// param1 = obj
function TLuaHttpExpr.Find_FolderClose(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFolder;
  i:Integer;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFolder(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if fobj<>nil then begin
    i:=FObjList.IndexOf(fobj);
    if i<>-1 then
      FObjList.Delete(i);
  end;
  Result:=0;
end;

// param1 = filename, param2 = flag, ret = obj
function TLuaHttpExpr.Find_FileFirst(State: TLuaState): Integer;
var
  arg:Integer;
  fname:string;
  attr:longint;
  fobj:TLuaFindFile;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  if arg>1 then
    attr:=lua_tointeger(State,2)
    else
      attr:=faAnyFile;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    try
      fobj:=TLuaFindFile.Create(fname,attr);
      try
        lua_pushlightuserdata(State,Pointer(fobj));
        FObjList.Add(fobj);
      except
        on e:Exception do begin
          fobj.Free;
          lastError:=3;
          lastErrMsg:=e.Message;
        end;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = obj, ret bool
function TLuaHttpExpr.Find_FileNext(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFile;
  bret:boolean;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFile(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if (fobj<>nil) and (FObjList.IndexOf(fobj)<>-1) then
    bret:=fobj.Next
    else
      bret:=False;
  Result:=1;
  lua_pushboolean(State,bret);
end;

function TLuaHttpExpr.Find_FileName(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFile;
  sret:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFile(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if (fobj<>nil) and (FObjList.IndexOf(fobj)<>-1) then begin
    sret:=fobj.GetFullName;
    if FUTF8File then
      sret:=AnsiToUtf8(sret);
    end else
      sret:='';
  Result:=1;
  lua_pushstring(State,sret);
end;

function TLuaHttpExpr.Find_FileAttr(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFile;
  rattr:Integer;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFile(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if (fobj<>nil) and (FObjList.IndexOf(fobj)<>-1) then
    rattr:=fobj.GetAttr
    else
      rattr:=0;
  Result:=1;
  lua_pushinteger(State,rattr);
end;

// param1 = obj
function TLuaHttpExpr.Find_FileClose(State: TLuaState): Integer;
var
  arg:Integer;
  fobj:TLuaFindFile;
  i:Integer;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fobj:=TLuaFindFile(lua_touserdata(State,1))
    else
      fobj:=nil;
  lua_pop(State,arg);
  if fobj<>nil then begin
    i:=FObjList.IndexOf(fobj);
    if i<>-1 then
      FObjList.Delete(i);
  end;
  Result:=0;
end;


// param1 = expression, param2 = flags, ret = RegEx
{
  flags
  define "(?ui)" is more easy in expression

  brrefDELIMITERS=1 shl 0;
  brrefBACKTRACKING=1 shl 1;
  brrefFREESPACING=1 shl 2;
  brrefIGNORECASE=1 shl 3;
  brrefSINGLELINE=1 shl 4;
  brrefMULTILINE=1 shl 5;
  brrefLATIN1=1 shl 6;
  brrefUTF8=1 shl 7;
  brrefUTF8CODEUNITS=1 shl 8;
  brrefAUTO=1 shl 11;
  brrefWHOLEONLY=1 shl 12;
  brrefCOMBINEOPTIMIZATION=1 shl 13;
}
function TLuaHttpExpr.RegEx_New(State: TLuaState): Integer;
var
  arg:Integer;
  exp:TBRRERegExp;
  expstr:string;
  flag:LongInt;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    expstr:=lua_tostring(State,1)
    else
      expstr:='';
  if arg>1 then
    flag:=lua_tointeger(State,2)
    else
      flag:=0;
  lua_pop(State,arg);
  if expstr<>'' then begin
    try
      exp:=TBRRERegExp.Create(expstr,flag);
      try
        lua_pushlightuserdata(State,Pointer(exp));
        FObjList.Add(exp);
      except
        on e:Exception do begin
          exp.Free;
          lastError:=3;
          lastErrMsg:=e.Message;
        end;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else begin
    lastError:=1;
    lastErrMsg:=rsInvalidRegexExpr;
  end;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = regex
function TLuaHttpExpr.RegEx_Delete(State: TLuaState): Integer;
var
  ix,arg:Integer;
  RegEx:TBRRERegExp;
begin
  ResetError;
  Result:=0;
  arg:=lua_gettop(State);
  if arg>0 then
    RegEx:=TBRRERegExp(lua_touserdata(State,1))
    else
      RegEx:=nil;
  lua_pop(State,arg);
  if RegEx<>nil then begin
    ix:=FObjList.IndexOf(RegEx);
    if ix<>-1 then
      FObjList.Delete(ix);
  end;
end;

// param1 = regex, param2 = source, param3= pos, ret = string, pos, len
function TLuaHttpExpr.RegEx_Match(State: TLuaState): Integer;
var
  arg:Integer;
  regex:TBRRERegExp;
  matchp:TBRRERegExpCaptures;
  src:string;
  ret:string;
  pos,len,utf8src:Integer;
begin
  ResetError;
  Result:=3;
  len:=0;
  arg:=lua_gettop(State);
  if arg>1 then begin
    regex:=TBRRERegExp(lua_touserdata(State,1));
    src:=lua_tostring(State,2);
    if arg>2 then
      pos:=lua_tointeger(State,3)
      else
        pos:=1;
  end else
    regex:=nil;
  if arg>3 then
    utf8src:=lua_tointeger(State,4)
    else
      utf8src:=-1;
  lua_pop(State,arg);
  ret:='';
  if (regex<>nil) and (FObjList.IndexOf(regex)<>-1) then begin
    try
     SetLength(matchp,0);
     if regex.Match(src,pos-1,pos,matchp,utf8src) then begin
       len:=matchp[0].EndCodePoint-matchp[0].StartCodePoint;
       ret:=Copy(src,matchp[0].StartCodePoint+1,len);
       pos:=matchp[0].StartCodePoint+1;
     end else
       pos:=0;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
        pos:=0;
        ret:='';
      end;
    end;
  end else begin
    lastError:=1;
    lastErrMsg:=rsInvalidRegEx;
  end;
  lua_pushstring(State,ret);
  lua_pushinteger(State,pos);
  lua_pushinteger(State,len);
end;

// param1 = regex, param2 = src, ret = table
function TLuaHttpExpr.RegEx_MatchAll(State: TLuaState): Integer;
var
  arg:Integer;
  regex:TBRRERegExp;
  src,temp:string;
  res:TBRRERegExpMultipleCaptures;
  row:TBRRERegExpCaptures;
  count,col,lcol,i,j,len,utf8src,rrow:Integer;
begin
  ResetError;
  Result:=1;
  arg:=lua_gettop(State);
  if arg>1 then begin
    regex:=TBRRERegExp(lua_touserdata(State,1));
    src:=lua_tostring(State,2);
  end else
    regex:=nil;
  if arg>2 then
    utf8src:=lua_tointeger(State,3)
    else
      utf8src:=-1;
  lua_pop(State,arg);
  if (regex<>nil) and (FObjList.IndexOf(regex)<>-1) then begin
    try
      SetLength(res,0);
      regex.MatchAll(src,res,utf8src);
      count:=Length(res);
      if count>0 then begin
        lua_newtable(State);
        lcol:=-1;
        rrow:=0;
        for j:=0 to count-1 do begin
          row:=res[j];
          col:=Length(row);
          if col>lcol then
            lcol:=col;
          if col>0 then begin
            Inc(rrow);
            for i:=0 to col-1 do
              if length(row)>0 then begin
                len:=row[i].EndCodePoint-row[i].StartCodePoint;
                temp:=Copy(src,row[i].StartCodePoint+1,len);
                // row,col=value
                lua_pushstring(State,Format('%d,%d',[j,i]));
                lua_pushstring(State,temp);
                lua_settable(State,-3);
              end;
          end;
        end;
        lua_pushstring(State,'Row');
        lua_pushinteger(State,rrow);
        lua_settable(State,-3);
        lua_pushstring(State,'Col');
        lua_pushinteger(State,lcol);
        lua_settable(State,-3);
      end else
        lua_pushnil(State);
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else begin
    lastError:=1;
    lastErrMsg:=rsInvalidRegEx;
  end;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = regex, param2 = src, ret = CSV string
function TLuaHttpExpr.RegEx_MatchAllCSV(State: TLuaState): Integer;
var
  arg:Integer;
  regex:TBRRERegExp;
  src,temp,csvres:string;
  res:TBRRERegExpMultipleCaptures;
  row:TBRRERegExpCaptures;
  count,col,i,j,len,utf8src:Integer;
  doc:TCSVBuilder;
  istr:TStringStream;
begin
  ResetError;
  Result:=1;
  arg:=lua_gettop(State);
  if arg>1 then begin
    regex:=TBRRERegExp(lua_touserdata(State,1));
    src:=lua_tostring(State,2);
  end else
    regex:=nil;
  if arg>2 then
    utf8src:=lua_tointeger(State,3)
    else begin
      if FuseUTF8String then
         utf8src:=-1
         else
          utf8src:=0;
    end;
  lua_pop(State,arg);
  csvres:='';
  if (regex<>nil) and (FObjList.IndexOf(regex)<>-1) then begin
    try
      SetLength(res,0);
      regex.MatchAll(src,res,utf8src);
      count:=Length(res);
      if count>0 then begin
        doc:=TCSVBuilder.Create;
        try
          doc.Delimiter:=CSVDelimit;
          istr:=TStringStream.Create('');
          try
            doc.SetOutput(istr);
            for j:=0 to count-1 do begin
              row:=res[j];
              col:=Length(row);
              if col>0 then begin
                for i:=0 to col-1 do begin
                  if Length(row)>0 then begin
                    len:=row[i].EndCodePoint-row[i].StartCodePoint;
                    temp:=Copy(src,row[i].StartCodePoint+1,len);
                    if utf8src<>-1 then
                      temp:=Utf8ToAnsi(temp);
                    doc.AppendCell(temp);
                  end;
                end;
                if j<>count-1 then
                  doc.AppendRow;
              end;
            end;
            csvres:=istr.DataString;
          finally
            istr.Free;
          end;
        finally
          doc.Free;
        end;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else begin
    lastError:=1;
    lastErrMsg:=rsInvalidRegEx;
  end;
  Result:=1;
  lua_pushstring(State,csvres);
end;

// param1 = regex, param2 = src, param3 = replace str, ret = str
function TLuaHttpExpr.RegEx_Replace(State: TLuaState): Integer;
var
  arg,utf8src:Integer;
  regex:TBRRERegExp;
  src,repstr,ret:string;
begin
  ResetError;
  Result:=1;
  arg:=lua_gettop(State);
  if arg>2 then begin
    regex:=TBRRERegExp(lua_touserdata(State,1));
    src:=lua_tostring(State,2);
    repstr:=lua_tostring(State,3);
  end else
    regex:=nil;
  if arg>3 then
    utf8src:=lua_tointeger(State,4)
    else
      utf8src:=-1;
  lua_pop(State,arg);
  if (regex<>nil) and (FObjList.IndexOf(regex)<>-1) then begin
    try
      ret:=regex.Replace(src,repstr,utf8src);
      lua_pushstring(State,ret);
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else begin
    lastError:=1;
    lastErrMsg:=rsInvalidRegEx;
  end;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = string, ret = string
function TLuaHttpExpr.HWP_ReadText(State: TLuaState): Integer;
var
  arg:Integer;
  fname,txt:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    try
      txt:=system.UTF8Encode(ReadHWPText(fname));
      lua_pushstring(State,txt);
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = url|file|string, param2 = 0|1|2 , ret = table
function TLuaHttpExpr.RSS_Read(State: TLuaState): Integer;
var
  arg,loc,i:Integer;
  url:string;
  rss:TRSSReader;
  item:TRSSItem;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    url:=lua_tostring(State,1)
    else
      url:='';
  if arg>1 then
    loc:=lua_tointeger(State,2)
    else
      loc:=0;
  lua_pop(State,arg);
  if url<>'' then begin
    try
      rss:=TRSSReader.Create;
      try
        rss.UTF8:=True;
        case loc of
        1: begin
             if FUTF8File then
               url:=Utf8ToAnsi(url);
             rss.LoadFromFile(url);
           end;
        2: rss.LoadFromString(url);
        else
          rss.LoadFromHttp(url);
        end;
        i:=0;
        lua_newtable(State);
        if rss.Count>0 then
          for item in rss.Items do begin
            lua_pushstring(State,'Title'+IntToStr(i));
            lua_pushstring(State,item.Title);
            lua_settable(State,-3);
            lua_pushstring(State,'Link'+IntToStr(i));
            lua_pushstring(State,item.Link);
            lua_settable(State,-3);
            lua_pushstring(State,'Date'+IntToStr(i));
            lua_pushstring(State,item.PubDate);
            lua_settable(State,-3);
            lua_pushstring(State,'Content'+IntToStr(i));
            lua_pushstring(State,item.Content);
            lua_settable(State,-3);
            lua_pushstring(State,'Desc'+IntToStr(i));
            lua_pushstring(State,item.Description);
            lua_settable(State,-3);
            lua_pushstring(State,'Author'+IntToStr(i));
            lua_pushstring(State,item.Author);
            lua_settable(State,-3);
            lua_pushstring(State,'Comments'+IntToStr(i));
            lua_pushstring(State,item.Comments);
            lua_settable(State,-3);
            lua_pushstring(State,'IsPermaLink'+IntToStr(i));
            lua_pushboolean(State,item.IsPermaLink);
            lua_settable(State,-3);
            Inc(i);
          end;
        lua_pushstring(State,'Docs');
        lua_pushstring(State,rss.Docs);
        lua_settable(State,-3);

        lua_pushstring(State,'Webmaster');
        lua_pushstring(State,rss.WebMaster);
        lua_settable(State,-3);

        lua_pushstring(State,'Copyright');
        lua_pushstring(State,rss.Copyright);
        lua_settable(State,-3);

        lua_pushstring(State,'Count');
        lua_pushinteger(State,rss.Count);
        lua_settable(State,-3);
      finally
        rss.Free;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State);
end;

// param1 = col, param2 = row; ret(no input) = col, row
function TLuaHttpExpr.Grid_ColRow(State: TLuaState): Integer;
var
  arg,col,row:Integer;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    col:=lua_tointeger(State,1)+1
    else
      col:=-1;
  if arg>1 then
    row:=lua_tointeger(State,2)+1
    else
      row:=-1;
  lua_pop(State,arg);
  with txtworker_main.FormLua do begin
    if col>0 then begin
      WorkGrid.ColCount:=col;
      SetColHeaders;
    end;
    if row>0 then begin
      WorkGrid.RowCount:=row;
      SetRowHeaders;
    end;
    if arg=0 then begin
      Result:=2;
      lua_pushinteger(State,WorkGrid.ColCount-1);
      lua_pushinteger(State,WorkGrid.RowCount-1);
    end else
      Result:=0;
  end;
end;

// param1 = col, param2 = row, param3 = str; ret(no input) str
function TLuaHttpExpr.Grid_Value(State: TLuaState): Integer;
var
  arg,col,row:Integer;
  str:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    col:=lua_tointeger(State,1)+1;
  if arg>1 then
    row:=lua_tointeger(State,2)+1;
  if arg>2 then
    str:=lua_tostring(State,3);
  lua_pop(State,arg);
  with txtworker_main.FormLua do begin
    try
      if arg>2 then begin
        StrGrid_SetValue(col, row, str);
        FHitWorkCell:=True;
        Result:=0;
      end else if arg>1 then begin
        if (col<WorkGrid.ColCount) and (row<WorkGrid.RowCount) then
          str:=WorkGrid.Cells[col,row]
          else
            str:='';
        Result:=1;
        lua_pushstring(State,str);
      end else begin
        lastError:=1;
        lastErrMsg:=format(rsSInvalidPara, ['Grid_Value']);
        Result:=0;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end;
end;

function TLuaHttpExpr.Grid_AutoColumn(State: TLuaState): Integer;
var
  arg,col:Integer;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    col:=lua_tointeger(State,1)
    else
      col:=-1;
  lua_pop(State,arg);
  if col>-1 then
     FormLua.WorkGrid.AutoSizeColumn(col)
     else begin
       FormLua.WorkGrid.AutoSizeColumns;
       FHitWorkCell:=False;
     end;
  Result:=0;
end;

// no param
function TLuaHttpExpr.Grid_Clear(State: TLuaState): Integer;
begin
  lua_pop(State,lua_gettop(State));
  StrGrid_clear;
  Result:=0;
end;

// param1 = filename, param2 = csv, param3 = delim, ret = boolean;
function TLuaHttpExpr.Grid_Load(State: TLuaState): Integer;
var
  arg:Integer;
  fname,delim:string;
  bcsv:Boolean;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  if arg>1 then
    bcsv:=lua_toboolean(State,2)
    else
      bcsv:=True;
  if arg>2 then
    delim:=lua_tostring(State,3)
    else
      delim:=CSVDelimit;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    try
      StrGrid_clear;
      if bcsv then begin
        //txtworker_main.FormLua.WorkGrid.LoadFromCSVFile(fname,delim[1],False)
        LoadCSVFile(fname,delim[1]);
        FLastCol:=FormLua.WorkGrid.ColCount-1;
        FLastRow:=FormLua.WorkGrid.RowCount-1;
        end else
          txtworker_main.FormLua.WorkGrid.LoadFromFile(fname);
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

function TLuaHttpExpr.Grid_LoadStr(State: TLuaState): Integer;
var
  arg:Integer;
  datastr,delim:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    datastr:=lua_tostring(State,1)
    else
      datastr:='';
  if arg>1 then
    delim:=lua_tostring(State,2)
    else
      delim:=CSVDelimit;
  lua_pop(State,arg);
  if datastr<>'' then begin
    try
      if delim='' then
        delim:=CSVDelimit;
      StrGrid_clear;
      LoadCSVStr(datastr,delim[1]);
      FLastCol:=FormLua.WorkGrid.ColCount-1;
      FLastRow:=FormLua.WorkGrid.RowCount-1;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// param1 = filename, param2 = csv boolean, param3 = delimiter, ret = boolean
function TLuaHttpExpr.Grid_Save(State: TLuaState): Integer;
var
  arg:Integer;
  fname,delim:string;
  bcsv:Boolean;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  if arg>1 then
    bcsv:=lua_toboolean(State,2)
    else
      bcsv:=True;
  if arg>2 then
    delim:=lua_tostring(State,3)
    else
      delim:=CSVDelimit;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    try
      if bcsv then
        //txtworker_main.FormLua.WorkGrid.SaveToCSVFile(fname,delim[1],False)
        SaveCSVFile(fname,delim[1])
        else
          txtworker_main.FormLua.WorkGrid.SaveToFile(fname);
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

function TLuaHttpExpr.Grid_SaveStr(State: TLuaState): Integer;
var
  arg:Integer;
  datastr,delim:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    delim:=lua_tostring(State,1)
    else
      delim:=CSVDelimit;
  lua_pop(State,arg);
  try
    datastr:=SaveCSVStr(delim[1])
  except
    on e:Exception do begin
      lastError:=2;
      lastErrMsg:=e.Message;
    end;
  end;
  Result:=1;
  lua_pushstring(State,datastr);
end;

// ret = table
function TLuaHttpExpr.Grid_ToTable(State: TLuaState): Integer;
var
  temp:string;
  row,col,j,i:Integer;
begin
  lua_pop(State,lua_gettop(State));
  row:=txtworker_main.FormLua.WorkGrid.RowCount-1;
  col:=txtworker_main.FormLua.WorkGrid.ColCount-1;
  Result:=1;
  if (row>0) and (col>0) then begin
    lua_newtable(State);
    lua_pushstring(State,'Col');
    lua_pushinteger(State,col);
    lua_settable(State,-3);
    lua_pushstring(State,'Row');
    lua_pushinteger(State,row);
    lua_settable(State,-3);
    for j:=0 to row-1 do
      for i:=0 to col-1 do begin
        temp:=txtworker_main.FormLua.WorkGrid.Cells[i+1,j+1];
        lua_pushstring(State,Format('%d,%d',[j,i]));
        lua_pushstring(State,temp);
        lua_settable(State,-3);
      end;
  end else
    lua_pushnil(State);
end;

// param1 = table, clear grid.
function TLuaHttpExpr.Grid_FromTable(State: TLuaState): Integer;
var
  arg,Keyidx,col:Integer;
  temp:string;
  tempnum:Double;
  keycol:array of string;
  keypos:array of Integer;
  function getkeyidx(const key:string):Integer;
  var
    i,l:Integer;
  begin
    l:=length(keycol)-1;
    Result:=-1;
    for i:=0 to l do
      if key=keycol[i] then begin
        Result:=i;
        Inc(keypos[i],1);
        break;
      end;
    if Result=-1 then begin
      SetLength(keycol,l+2);
      keycol[l+1]:=key;
      Result:=l+1;
      SetLength(keypos,l+2);
      keypos[Result]:=2;
      StrGrid_SetValue(1,Result,key,False);
    end;
  end;

begin
  arg:=lua_gettop(State);
  if (arg=1) and lua_istable(State,arg) then begin
    Inc(dumplvl);
    try
      SetLength(keycol,1);
      SetLength(keypos,1);
      keycol[0]:='';
      keypos[0]:=1;
      Keyidx:=1;
      StrGrid_clear;
      lua_pushnil(State);
      while lua_next(State,arg)<>0 do begin
        // key
        if lua_isstring(State,-2) then
          temp:=lua_tostring(State,-2)
            else if lua_isnumber(State,-2) then begin
              tempnum:=lua_tonumber(State,-2);
              temp:='#'+NumToStr(tempnum);
              end else
                temp:='';
        Keyidx:=getkeyidx(temp);
        col:=keypos[Keyidx];
        // value
        if Keyidx>0 then begin
          if lua_istable(State,-1) then begin
            Grid_FromTable(State);
          end else if lua_isnumber(State,-1) then begin
            tempnum:=lua_tonumber(State,-1);
            temp:=NumToStr(tempnum);
            StrGrid_SetValue(col,Keyidx,temp,False);
          end else if lua_isstring(State,-1) then begin
            temp:=lua_tostring(State,-1);
            StrGrid_SetValue(col,Keyidx,temp,False);
          end else if lua_isboolean(State,-1) then begin
            if lua_toboolean(State,-1) then
              temp:='true'
              else
                temp:='false';
            StrGrid_SetValue(col,Keyidx,temp,False);
          end else begin
            temp:='unknown';
            StrGrid_SetValue(col,Keyidx,temp,False);
          end;
        end;
        lua_pop(State,1);
      end;
      lua_pop(State,1);
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
        if dumplvl>1 then begin
          Dec(dumplvl);
          raise;
        end;
      end;
    end;
    Dec(dumplvl);
    SetLength(keypos,0);
    SetLength(keycol,0);
    SetColHeaders;
    SetRowHeaders;
    FormLua.WorkGrid.AutoSizeColumns;
  end else
    lua_pop(State,arg);
  Result:=0;
end;

// param1 = integer
function TLuaHttpExpr.Grid_DeleteRow(State: TLuaState): Integer;
var
  arg,i:Integer;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    i:=lua_tointeger(State,1)+1
    else
      i:=-1;
  lua_pop(State,arg);
  if (i>0) and (i<txtworker_main.FormLua.WorkGrid.RowCount) then begin
    txtworker_main.FormLua.WorkGrid.DeleteRow(i);
    FLastRow:=i;
    SetRowHeaders;
  end;
  Result:=0;
end;

// param1 = integer
function TLuaHttpExpr.Grid_DeleteCol(State: TLuaState): Integer;
var
  arg,i:Integer;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    i:=lua_tointeger(State,1)+1
    else
      i:=-1;
  lua_pop(State,arg);
  if (i>0) and (i<txtworker_main.FormLua.WorkGrid.ColCount) then begin
    txtworker_main.FormLua.WorkGrid.DeleteCol(i);
    FLastCol:=i;
    SetColHeaders;
  end;
  Result:=0;
end;

const
  xlsExt:array[0..2] of string=(
    STR_EXCEL_EXTENSION,
    STR_OPENDOCUMENT_CALC_EXTENSION,
    STR_OOXML_EXCEL_EXTENSION
    );

function FindExcelFile(var fname:string):boolean;
var
  i:Integer;
begin
  Result:=FileExists(fname);
  if not Result then
    for i:=Low(xlsExt) to high(xlsExt) do begin
      fname:=ChangeFileExt(fname,xlsExt[i]);
      if FileExists(fname) then
        Result:=True;
        break;
    end;
end;

// param1 = filename, param2 = sheetindex
function TLuaHttpExpr.Grid_LoadExcel(State: TLuaState): Integer;
var
  arg:Integer;
  fname,temp:string;
  sindex,i,j,k:Integer;
  worksht:TsWorksheet;
  workbk:TsWorkbook;
  lCell:PCell;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  if arg>1 then
    sindex:=lua_tointeger(State,2)
    else
      sindex:=0;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    with txtworker_main.FormLua do begin
      StrGrid_clear;
      try
        workbk:=TsWorkbook.Create;
        try
          if FindExcelFile(fname) then begin
            workbk.ReadFromFile(fname);
            worksht:=workbk.GetWorksheetByIndex(sindex);
            lCell:=worksht.GetFirstCell();
            for k:=0 to worksht.GetCellCount-1 do begin
              i:=lCell^.Col;
              j:=lCell^.Row;
              temp:=worksht.ReadAsUTF8Text(j,i);
              StrGrid_SetValue(i+1,j+1,temp,False);
              lCell:=worksht.GetNextCell();
            end;
          end else
            lastError:=2;
        finally
          workbk.Free;
        end;

      except
        on e:Exception do begin
          lastError:=3;
          lastErrMsg:=e.Message;
        end;
      end;
      SetColHeaders;
      SetRowHeaders;
      WorkGrid.AutoSizeColumns;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// param1 = filename, type : 0 - excel8, 1 - opendoc, 2 - ooxml
function TLuaHttpExpr.Grid_SaveExcel(State: TLuaState): Integer;
var
  arg:Integer;
  fname,temp:string;
  workbk:TsWorkbook;
  worksht:TsWorksheet;
  i,j:Integer;
  shtype:TsSpreadsheetFormat;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    fname:=lua_tostring(State,1)
    else
      fname:='';
  if arg>1 then begin
    case lua_tointeger(State,2) of
    1:shtype:=sfOpenDocument;
    2:shtype:=sfOOXML;
    else
      shtype:=sfExcel8;
    end;
  end else
    shtype:=sfExcel8;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then
      fname:=Utf8ToAnsi(fname);
    with txtworker_main.FormLua do begin
      try
        workbk:=TsWorkbook.Create;
        try
          worksht:=workbk.AddWorksheet('sheet1');
          for j:=1 to WorkGrid.RowCount-1 do
            for i:=1 to WorkGrid.ColCount-1 do begin
              temp:=WorkGrid.Cells[i,j];
              if temp<>'' then
                worksht.WriteUTF8Text(j-1,i-1,temp);
            end;
          // change extension
          case shtype of
          sfExcel8:temp:=STR_EXCEL_EXTENSION;
          sfOOXML:temp:=STR_OOXML_EXCEL_EXTENSION;
          sfOpenDocument:temp:=STR_OPENDOCUMENT_CALC_EXTENSION;
          end;
          fname:=ChangeFileExt(fname,temp);
          workbk.WriteToFile(fname,shtype,True);
        finally
          workbk.Free;
        end;
      except
        on e:Exception do begin
          lastError:=2;
          lastErrMsg:=e.Message;
        end;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// sort column
function TLuaHttpExpr.Grid_SortCol(State: TLuaState): Integer;
var
  arg:Integer;
  index:Integer;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    index:=lua_tointeger(State,1)+1
    else
      index:=-1;
  lua_pop(State,arg);
  if index>-1 then begin
    FormLua.WorkGrid.SortColRow(True,index);
  end else
    lastError:=1;
  Result:=0;
end;

function TLuaHttpExpr.CSVDelimiter(State: TLuaState): Integer;
var
  arg:Integer;
  str:string;
begin
  str:='';
  arg:=lua_gettop(State);
  if arg>0 then
    if lua_isstring(State,1) then
      str:=lua_tostring(State,1);
  lua_pop(State,arg);
  if Length(str)>0 then begin
    CSVDelimit:=str[1];
    Result:=0;
  end else begin
    str:=CSVDelimit;
    lua_pushstring(State,str);
    Result:=1;
  end;
end;


// param1 = msg, param2 = cap, ret=str
function TLuaHttpExpr.InputBox(State: TLuaState): Integer;
var
  arg:Integer;
  rstr,msg,cap:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    msg:=lua_tostring(State,1)
    else
      msg:='';
  if arg>1 then
    cap:=lua_tostring(State,2)
    else
      msg:='';
  lua_pop(State,arg);
  rstr:=Dialogs.InputBox(cap,msg,'');
  Result:=1;
  lua_pushstring(State,rstr);
end;

function DefMessageDlg(const msg: string; const cap: string; DlgType: TMsgDlgType;
  Buttons: TMsgDlgButtons; DefButton: Integer): Integer;
var
  i: Integer;
  btn: TBitBtn;
begin
  with CreateMessageDialog(msg, DlgType, Buttons) do
  try
    Position:=poOwnerFormCenter;
    Caption := cap;
    for i := 0 to ControlCount - 1 do
    begin
      if Controls[i] is TBitBtn then
      begin
        btn := TBitBtn(Controls[i]);
        btn.Default := btn.ModalResult = DefButton;
        if btn.Default then
          ActiveControl := Btn;
      end;
    end;
    Result := ShowModal;
  finally
    Free;
  end;
end;

// param1=str, param2=cap, param3=defbtn, ret=bool
function TLuaHttpExpr.QueryBoxYesNo(State: TLuaState): Integer;
var
  arg,defbtn:Integer;
  msg,cap:string;
  rbox:Boolean;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    msg:=lua_tostring(State,1)
    else
      msg:='';
  if arg>1 then
    cap:=lua_tostring(State,2)
    else
      cap:='txtwoker';
  if arg>2 then
    case lua_tointeger(State,3) of
    1: defbtn:=mrNo;
    else
      defbtn:=mrYes;
    end
    else
      defbtn:=mrYes;
  lua_pop(State,arg);
  rbox:=DefMessageDlg(msg,cap,mtConfirmation,mbYesNo,defbtn)=mrYes;
  Result:=1;
  lua_pushboolean(State,rbox);
end;

// param1=str, param2=cap, param3=defbtn, ret=0,1,2
function TLuaHttpExpr.QueryBoxYesNoCancel(State: TLuaState): Integer;
var
  arg,defbtn:Integer;
  msg,cap:string;
begin
  arg:=lua_gettop(State);
  if arg>0 then
    msg:=lua_tostring(State,1)
    else
      msg:='';
  if arg>1 then
    cap:=lua_tostring(State,2)
    else
      cap:='txtwoker';
  if arg>2 then
    case lua_tointeger(State,3) of
    1: defbtn:=mrNo;
    2: defbtn:=mrCancel;
    else
      defbtn:=mrYes;
    end
    else
      defbtn:=mrYes;
  lua_pop(State,arg);
  defbtn:=DefMessageDlg(msg,cap,mtConfirmation,mbYesNoCancel,defbtn);
  Result:=1;
  if defbtn=mrYes then
    defbtn:=0
    else if defbtn=mrCancel then
      defbtn:=2
      else
        defbtn:=1;
  lua_pushinteger(State,defbtn);
end;

// filename, basedir, files, password, subdir, ret = bool
function TLuaHttpExpr.Zip_Add(State: TLuaState): Integer;
var
  arg,subdir:Integer;
  fname,bdir,files,password:string;
  zip:TAbZipper;
  sopt:TAbStoreOptions;
begin
  arg:=lua_gettop(State);
  if arg>2 then begin
    fname:=lua_tostring(State,1);
    bdir:=lua_tostring(State,2);
    files:=lua_tostring(State,3);
  end else
    fname:='';
  if arg>3 then
    password:=lua_tostring(State,4)
    else
      password:='';
  if arg>4 then
    subdir:=lua_tointeger(State,5)
    else
      subdir:=1;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then begin
      fname:=Utf8ToAnsi(fname);
      bdir:=Utf8ToAnsi(bdir);
      files:=Utf8ToAnsi(files);
    end;
    try
      zip:=TAbZipper.Create(nil);
      try
        sopt:=[soStripDrive,soRemoveDots];
        if subdir<>0 then
          sopt:=sopt+[soRecurse];
        zip.AutoSave:=True;
        zip.StoreOptions:=sopt;
        zip.Logging:=False;
        zip.DOSMode:=False;
        zip.CompressionMethodToUse:=smBestMethod;
        zip.BaseDirectory:=bdir;
        zip.FileName:=fname;
        zip.Password:=password;
        zip.AddFiles(files,faAnyFile);
      finally
        zip.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// filename, basedir, name, text, password, subdir, ret = bool
function TLuaHttpExpr.Zip_AddText(State: TLuaState): Integer;
var
  arg,subdir:Integer;
  fname,bdir,files,password,txt:string;
  zip:TAbZipper;
  sopt:TAbStoreOptions;
  tstm:TStringStream;
begin
  arg:=lua_gettop(State);
  if arg>3 then begin
    fname:=lua_tostring(State,1);
    bdir:=lua_tostring(State,2);
    files:=lua_tostring(State,3);
    txt:=lua_tostring(State,4);
  end else
    fname:='';
  if arg>4 then
    password:=lua_tostring(State,5)
    else
      password:='';
  if arg>5 then
    subdir:=lua_tointeger(State,6)
    else
      subdir:=1;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then begin
      fname:=Utf8ToAnsi(fname);
      bdir:=Utf8ToAnsi(bdir);
      files:=Utf8ToAnsi(files);
    end;
    try
      tstm:=TStringStream.Create(txt);
      try
        zip:=TAbZipper.Create(nil);
        try
          sopt:=[soStripDrive,soRemoveDots];
          if subdir<>0 then
            sopt:=sopt+[soRecurse];
          zip.AutoSave:=True;
          zip.StoreOptions:=sopt;
          zip.Logging:=False;
          zip.DOSMode:=False;
          zip.CompressionMethodToUse:=smBestMethod;
          zip.BaseDirectory:=bdir;
          zip.FileName:=fname;
          zip.Password:=password;
          zip.AddFromStream(files,tstm);
        finally
          zip.Free;
        end;
      finally
        tstm.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// filename, basedir, files, password, subdir, ret = bool
function TLuaHttpExpr.Zip_Delete(State: TLuaState): Integer;
var
  arg,subdir:Integer;
  fname,bdir,files,password:string;
  zip:TAbZipper;
  sopt:TAbStoreOptions;
begin
  arg:=lua_gettop(State);
  if arg>2 then begin
    fname:=lua_tostring(State,1);
    bdir:=lua_tostring(State,2);
    files:=lua_tostring(State,3);
  end else
    fname:='';
  if arg>3 then
    password:=lua_tostring(State,4)
    else
      password:='';
  if arg>4 then
    subdir:=lua_tointeger(State,5)
    else
      subdir:=1;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then begin
      fname:=Utf8ToAnsi(fname);
      bdir:=Utf8ToAnsi(bdir);
      files:=Utf8ToAnsi(files);
    end;
    try
      zip:=TAbZipper.Create(nil);
      try
        sopt:=[soStripDrive,soRemoveDots];
        if subdir<>0 then
          sopt:=sopt+[soRecurse];
        zip.AutoSave:=True;
        zip.StoreOptions:=sopt;
        zip.Logging:=False;
        zip.DOSMode:=False;
        zip.CompressionMethodToUse:=smBestMethod;
        zip.BaseDirectory:=bdir;
        zip.FileName:=fname;
        zip.Password:=password;
        zip.DeleteFiles(files);
      finally
        zip.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// filename, basedir, files, password, subdir, ret = bool
function TLuaHttpExpr.Zip_Freshen(State: TLuaState): Integer;
var
  arg,subdir:Integer;
  fname,bdir,files,password:string;
  zip:TAbZipper;
  sopt:TAbStoreOptions;
begin
  arg:=lua_gettop(State);
  if arg>2 then begin
    fname:=lua_tostring(State,1);
    bdir:=lua_tostring(State,2);
    files:=lua_tostring(State,3);
  end else
    fname:='';
  if arg>3 then
    password:=lua_tostring(State,4)
    else
      password:='';
  if arg>4 then
    subdir:=lua_tointeger(State,5)
    else
      subdir:=1;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then begin
      fname:=Utf8ToAnsi(fname);
      bdir:=Utf8ToAnsi(bdir);
      files:=Utf8ToAnsi(files);
    end;
    try
      zip:=TAbZipper.Create(nil);
      try
        sopt:=[soStripDrive,soRemoveDots];
        if subdir<>0 then
          sopt:=sopt+[soRecurse];
        zip.AutoSave:=True;
        zip.StoreOptions:=sopt;
        zip.Logging:=False;
        zip.DOSMode:=False;
        zip.CompressionMethodToUse:=smBestMethod;
        zip.BaseDirectory:=bdir;
        zip.FileName:=fname;
        zip.Password:=password;
        zip.FreshenFiles(files);
      finally
        zip.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// filename, basedir, files, password, subdir, ret = bool
function TLuaHttpExpr.Zip_Extract(State: TLuaState): Integer;
var
  arg,subdir:Integer;
  fname,bdir,files,password:string;
  unzip:TAbUnZipper;
  sopt:TAbExtractOptions;
begin
  arg:=lua_gettop(State);
  if arg>2 then begin
    fname:=lua_tostring(State,1);
    bdir:=lua_tostring(State,2);
    files:=lua_tostring(State,3);
  end else
    fname:='';
  if arg>3 then
    password:=lua_tostring(State,4)
    else
      password:='';
  if arg>4 then
    subdir:=lua_tointeger(State,5)
    else
      subdir:=1;
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then begin
      fname:=Utf8ToAnsi(fname);
      bdir:=Utf8ToAnsi(bdir);
      files:=Utf8ToAnsi(files);
    end;
    try
      unzip:=TAbUnZipper.Create(nil);
      try
        sopt:=[];
        if subdir<>0 then
          sopt:=sopt+[eoCreateDirs];
        unzip.ExtractOptions :=sopt;
        unzip.Logging:=False;
        unzip.BaseDirectory:=bdir;
        unzip.FileName:=fname;
        unzip.Password:=password;
        unzip.ExtractFiles(files);
      finally
        unzip.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  lua_pushboolean(State,lastError=0);
end;

// filename, basedir, name, password, ret = bool
function TLuaHttpExpr.Zip_ExtractText(State: TLuaState): Integer;
var
  arg,subdir:Integer;
  fname,bdir,files,password:string;
  unzip:TAbUnZipper;
  sopt:TAbExtractOptions;
  tstm:TStringStream;
begin
  arg:=lua_gettop(State);
  if arg>2 then begin
    fname:=lua_tostring(State,1);
    bdir:=lua_tostring(State,2);
    files:=lua_tostring(State,3);
  end else
    fname:='';
  if arg>3 then
    password:=lua_tostring(State,4)
    else
      password:='';
  lua_pop(State,arg);
  if fname<>'' then begin
    if FUTF8File then begin
      fname:=Utf8ToAnsi(fname);
      bdir:=Utf8ToAnsi(bdir);
      files:=Utf8ToAnsi(files);
    end;
    try
      tstm:=TStringStream.Create('');
        try
        unzip:=TAbUnZipper.Create(nil);
        try
          sopt:=[];
          unzip.ExtractOptions :=sopt;
          unzip.Logging:=False;
          unzip.BaseDirectory:=bdir;
          unzip.FileName:=fname;
          unzip.Password:=password;
          unzip.ExtractToStream(files,tstm);
          files:=tstm.DataString;
        finally
          unzip.Free;
        end;
      finally
        tstm.Free;
      end;
    except
      on e:Exception do begin
        lastError:=3;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State)
    else
      lua_pushstring(State,files);
end;

// ret, key value. no keypressed 0
function TLuaHttpExpr.Check_KeyPress(State: TLuaState): Integer;
var
  arg,waitkey:Integer;
  lastkey:Integer;
begin
  arg:=lua_gettop(State);
  if arg>0 then
     waitkey:=lua_tointeger(State,1)
     else
       waitkey:=0;
  lua_pop(State,arg);
  lastkey:=FormLua.lastkey;
  FormLua.lastkey:=0;
  Application.ProcessMessages;
  if waitkey<>0 then
    while lastkey=0 do begin
      lastkey:=FormLua.lastkey;
      FormLua.lastkey:=0;
      Application.ProcessMessages;
      Sleep(10);
    end;
  Result:=1;
  if lastkey=0 then
    lua_pushnil(State)
    else
      lua_pushinteger(State,lastkey);
end;

// txt = input formula, return = txt result
function TLuaHttpExpr.SolveFormula(State: TLuaState): Integer;
var
  arg:Integer;
  mysolver:TSimplefmParser;
  formula:string;
begin
  ResetError;
  arg:=lua_gettop(State);
  if arg>0 then
    formula:=lua_tostring(State,1)
    else
      formula:='';
  lua_pop(State,arg);
  if formula<>'' then begin
    try
      mysolver:=TSimplefmParser.Create;
      try
        mysolver.Precision:=80;
        mysolver.DoGenerateRPNQueue(formula);
        if 0<>mysolver.SolveRPNQueue(formula) then
          lastError:=1;
      finally
        mysolver.Free;
      end;
    except
      on e:Exception do begin
        lastError:=2;
        lastErrMsg:=e.Message;
      end;
    end;
  end else
    lastError:=1;
  Result:=1;
  if lastError<>0 then
    lua_pushnil(State)
    else
      lua_pushstring(State,formula);
end;

end.

