unit uluaFindFile;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  _MaxLuaFindFileDeep=120;

type

  { TLuaFindBase }

  TLuaFindBase=class
    protected
      FFileRec:TSearchRec;
      FindFlag:Integer;
      FCloseFree:Boolean;
      FPath,FFilename:string;
    public
      constructor Create(BaseDir: string; Attribute: Integer=faAnyFile);
      destructor Destroy; override;

      function Next:Boolean; virtual;
      function GetName:string;
      function GetFullName:string;
      function GetAttr:longint;
      function IsFolder:Boolean;
      function IsReadOnly:Boolean;
      function IsSysFile:Boolean;
      function IsHidden:Boolean;

      property Path:string read FPath;
      property Filename:string read FFilename;
  end;

  { TLuaFindFolder }

  TLuaFindFolder=class
    private
      FFindObj:array[0.._MaxLuaFindFileDeep] of TLuaFindBase;
      FFindDeep:Integer;
      FNewFolder:Boolean;
      FRecursive:Boolean;
    public
      constructor Create(BaseDir: string; Recursive: Boolean=True);
      destructor Destroy; override;

      function Next:Boolean;
      function GetName:string;
      function GetFullName:string;
      function GetAttr:longint;
      function IsFolder:Boolean;
      function IsReadOnly:Boolean;
      function IsSysFile:Boolean;
      function IsHidden:Boolean;
  end;

  { TLuaFindFile }

  TLuaFindFile=class(TLuaFindBase)
    public
      constructor Create(BaseDir: string; Attribute: Integer=faAnyFile);
      function Next: Boolean; override;
  end;

implementation

{ TLuaFindFolder }

constructor TLuaFindFolder.Create(BaseDir: string; Recursive:Boolean);
begin
  FRecursive:=Recursive;
  FFindDeep:=0;
  FFindObj[0]:=TLuaFindBase.Create(BaseDir,faAnyFile);
  FNewFolder:=False;
  if (FFindObj[0].IsFolder) and
    ((FFindObj[0].GetName='.') or (FFindObj[0].GetName='.')) then
      Next;
end;

destructor TLuaFindFolder.Destroy;
begin
  if FNewFolder then
    Inc(FFindDeep);
  while FFindDeep>0 do begin
    FFindObj[FFindDeep].Free;
    Dec(FFindDeep);
  end;
  inherited Destroy;
end;

function TLuaFindFolder.Next: Boolean;
var
  bDone:Boolean;
  name,newpath:string;
  i:Integer;
begin
  bDone:=False;
  Result:=False;
  repeat
    if FNewFolder or FFindObj[FFindDeep].Next then begin
      if FNewFolder then begin
        Inc(FFindDeep);
        FNewFolder:=False;
      end;
      if FFindObj[FFindDeep].IsFolder then begin
        name:=FFindObj[FFindDeep].GetName;
        if (name<>'.') and (name<>'..') then begin
          Result:=True;
          // prepare sub folders
          if FRecursive then
            if FFindDeep<64 then begin
              newpath:=ExcludeTrailingPathDelimiter(FFindObj[0].FPath);
              for i:=0 to FFindDeep do
                if newpath<>'' then
                  newpath:=newpath+PathDelim+FFindObj[i].GetName
                  else
                    newpath:=FFindObj[i].GetName;
              newpath:=newpath+PathDelim+FFindObj[0].FFilename;
              FFindObj[FFindDeep+1]:=TLuaFindBase.Create(newpath);
              FNewFolder:=FFindObj[FFindDeep+1].GetName<>'';
              if not FNewFolder then
                FFindObj[FFindDeep+1].Free;
            end;
        end;
      end;
    end else if FFindDeep>0 then begin
      FFindObj[FFindDeep].Free;
      Dec(FFindDeep);
    end else
      bDone:=True;
  until bDone or Result;
end;

function TLuaFindFolder.GetName: string;
begin
  Result:=FFindObj[FFindDeep].GetName;
end;

function TLuaFindFolder.GetFullName: string;
var
  i:Integer;
begin
  Result:=ExcludeTrailingPathDelimiter(FFindObj[0].FPath);
  for i:=0 to FFindDeep do
    if Result='' then
      Result:=FFindObj[i].GetName
      else
        Result:=Result+PathDelim+FFindObj[i].GetName;
end;

function TLuaFindFolder.GetAttr: longint;
begin
  Result:=FFindObj[FFindDeep].GetAttr;
end;

function TLuaFindFolder.IsFolder: Boolean;
begin
  Result:=FFindObj[FFindDeep].IsFolder;
end;

function TLuaFindFolder.IsReadOnly: Boolean;
begin
  Result:=FFindObj[FFindDeep].IsReadOnly;
end;

function TLuaFindFolder.IsSysFile: Boolean;
begin
  Result:=FFindObj[FFindDeep].IsSysFile;
end;

function TLuaFindFolder.IsHidden: Boolean;
begin
  Result:=FFindObj[FFindDeep].IsHidden;
end;

{ TLuaFindFile }

constructor TLuaFindFile.Create(BaseDir: string; Attribute: Integer);
begin
  inherited Create(BaseDir,Attribute);
  if IsFolder then
    Next;
end;

function TLuaFindFile.Next: Boolean;
begin
  Result:=False;
  if FindFlag=0 then begin
    while inherited Next do begin
      if not IsFolder then begin
        Result:=True;
        break;
      end;
    end;
  end;
end;

{ TLuaFindBase }

constructor TLuaFindBase.Create(BaseDir: string; Attribute: Integer=faAnyFile);
begin
  FPath:=ExtractFileDir(BaseDir);
  FFilename:=ExtractFileName(BaseDir);
  if FFilename='' then begin
    FFilename:='*.*';
    BaseDir:=FPath+PathDelim+FFilename;
  end;
  FindFlag:=FindFirst(BaseDir,Attribute,FFileRec);
  FCloseFree:=FindFlag=0;
end;

destructor TLuaFindBase.Destroy;
begin
  if FCloseFree then
    FindClose(FFileRec);
  inherited Destroy;
end;

function TLuaFindBase.Next: Boolean;
begin
  if FCloseFree then
    FindFlag:=FindNext(FFileRec)
    else
      FindFlag:=-1;
  Result:=FindFlag=0;
end;

function TLuaFindBase.GetName: string;
begin
  if FindFlag=0 then
    Result:=FFileRec.Name
    else
      Result:='';
end;

function TLuaFindBase.GetFullName: string;
begin
  if FindFlag=0 then begin
    if FPath<>'' then
      Result:=ExcludeTrailingPathDelimiter(FPath)+PathDelim+FFileRec.Name
      else
        Result:=FFileRec.Name;
    end else
      Result:='';
end;

function TLuaFindBase.GetAttr: longint;
begin
  if FindFlag=0 then
    Result:=FFileRec.Attr
    else
      Result:=0;
end;

function TLuaFindBase.IsFolder: Boolean;
begin
  if FindFlag=0 then
    Result:=(FFileRec.Attr and faDirectory)<>0
    else
      Result:=False;
end;

function TLuaFindBase.IsReadOnly: Boolean;
begin
  if FindFlag=0 then
    Result:=(FFileRec.Attr and faReadOnly)<>0
    else
      Result:=False;
end;

function TLuaFindBase.IsSysFile: Boolean;
begin
  if FindFlag=0 then
    Result:=(FFileRec.Attr and faSysFile)<>0
    else
      Result:=False;
end;

function TLuaFindBase.IsHidden: Boolean;
begin
  if FindFlag=0 then
    Result:=(FFileRec.Attr and faHidden)<>0
    else
      Result:=False;
end;

end.

