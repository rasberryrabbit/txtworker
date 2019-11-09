unit txtworker_main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, SynEdit, SynHighlighterAny, SynCompletion,
  SynGutterBase, SynGutterMarks, SynGutterLineNumber, SynGutterChanges,
  SynGutter, SynGutterCodeFolding, Forms, Controls, Graphics, Dialogs, Grids,
  Menus, ActnList, StdActns, StdCtrls, ExtCtrls, types, LCLType, OMultiPanel;

type

  { TFormLua }

  TFormLua = class(TForm)
    EditInsertPath: TAction;
    LuaReInstance: TAction;
    EditCodeComp1: TAction;
    FileOpen1: TAction;
    FileNew1: TAction;
    ActionInsRegexMatch: TAction;
    ActionInsRegex: TAction;
    EditRedo1: TAction;
    ActionImport: TAction;
    ActionExport: TAction;
    ActionTempHttp: TAction;
    ActionCopy: TAction;
    ActionRun: TAction;
    ActionList1: TActionList;
    EditCopy1: TEditCopy;
    EditCut1: TEditCut;
    EditDelete1: TEditDelete;
    EditPaste1: TEditPaste;
    EditSelectAll1: TEditSelectAll;
    EditUndo1: TEditUndo;
    FileExit1: TFileExit;
    FileSaveAs1: TFileSaveAs;
    MainMenu1: TMainMenu;
    Memo1: TMemo;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    OMultiPanel1: TOMultiPanel;
    OpenDialog1: TOpenDialog;
    FileOpenDlg: TOpenDialog;
    PopupMenuEdit: TPopupMenu;
    PopupMenuGrid: TPopupMenu;
    SaveDialog1: TSaveDialog;
    EditPathDlg: TSelectDirectoryDialog;
    SynEdit1: TSynEdit;
    SynCompletion1: TSynCompletion;
    WorkGrid: TStringGrid;
    procedure ActionCopyExecute(Sender: TObject);
    procedure ActionExportExecute(Sender: TObject);
    procedure ActionImportExecute(Sender: TObject);
    procedure ActionInsRegexExecute(Sender: TObject);
    procedure ActionInsRegexMatchExecute(Sender: TObject);
    procedure ActionRunExecute(Sender: TObject);
    procedure ActionTempHttpExecute(Sender: TObject);
    procedure EditCodeComp1Execute(Sender: TObject);
    procedure EditCopy1Execute(Sender: TObject);
    procedure EditCut1Execute(Sender: TObject);
    procedure EditDelete1Execute(Sender: TObject);
    procedure EditInsertPathExecute(Sender: TObject);
    procedure EditPaste1Execute(Sender: TObject);
    procedure EditPaste1Update(Sender: TObject);
    procedure EditRedo1Execute(Sender: TObject);
    procedure EditRedo1Update(Sender: TObject);
    procedure EditSelectAll1Execute(Sender: TObject);
    procedure EditUndo1Execute(Sender: TObject);
    procedure EditUndo1Update(Sender: TObject);
    procedure FileNew1Execute(Sender: TObject);
    procedure FileOpen1Execute(Sender: TObject);
    procedure FileSaveAs1Accept(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of String);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure LuaReInstanceExecute(Sender: TObject);
    procedure SynCompletion1CodeCompletion(var Value: string;
      SourceValue: string; var SourceStart, SourceEnd: TPoint;
      KeyChar: TUTF8Char; Shift: TShiftState);
    procedure SetCaption(str:string);
    procedure WorkGridCompareCells(Sender: TObject; ACol, ARow, BCol,
      BRow: Integer; var Result: integer);
  private
    procedure LoadPosition;
    procedure SavePosition;
    { private declarations }
  public
    lastkey:Integer;
    function CheckModified:boolean;
    { public declarations }
  end;

var
  FormLua: TFormLua;

implementation

uses lua,lua52,uLuaHttpExpr,LuaSyntax,DefaultTranslator,gettext,Translations,
  IniFiles;

var
  Worker:TLua;
  AppConf:String;

resourcestring
  rsStartScript = '***** Start Script';
  rsEndScript = '***** End Script';
  rsLostYourScri = 'Lost your script in Editor. Are you sure?';
  rsVregexRegEx_ = 'Vregex=RegEx_New(pattern);%sVtable=RegEx_MatchAll(Vregex,'
    +'text);%s--Dosomething...%sRegEx_Delete(Vregex);%s';
  rsVregexRegEx2_ = 'Vregex=RegEx_New(pattern);%sVtxt,Vpos,Vlen=RegEx_Match(Vregex,'
    +'text,1);%s--Dosomething...%sRegEx_Delete(Vregex);%s';
  rsScriptIsModi = 'Script is modified. Save changes?';
  rsLuaGridS = 'TxtWorker - %s';
  rsLuaInstanceR = '***** Lua Instance ReInitialized.';



{$R *.lfm}

{ TFormLua }

procedure TFormLua.FormCreate(Sender: TObject);
var
  LuaSyn:TSynLuaSyn;
  lng,lngf:string;
begin
  Worker:=TLuaHttpExpr.Create;
  LuaSyn:=TSynLuaSyn.Create(SynEdit1);
  SynEdit1.Highlighter:=LuaSyn;

  GetLanguageIDs(lng,lngf);
  Translations.TranslateUnitResourceStrings('LCLStrConsts', 'lclstrconsts.%s.po', lng, lngf);

  AppConf:=ChangeFileExt(GetUserDir+ExtractFileName(ParamStr(0)),'.ini');
  LoadPosition;
end;

// restore and save panel size
procedure TFormLua.LoadPosition;
var
  ini:TIniFile;
  w,h:Integer;
begin
  try
    ini:=TIniFile.Create(AppConf);
    try
      w:=ini.ReadInteger('Panel','Width',630);
      h:=ini.ReadInteger('Panel','Height',450);
      Width:=w;
      Height:=h;
      OMultiPanel1.LoadPositionsFromIniFile(ini,'Panel','Position');
    finally
      ini.Free;
    end;
  except
  end;
end;

procedure TFormLua.SavePosition;
var
  ini:TIniFile;
begin
  try
    ini:=TIniFile.Create(AppConf);
    try
      OMultiPanel1.SavePositionsToIniFile(ini,'Panel','Position');
      if WindowState<>wsMaximized then begin
        ini.WriteInteger('Panel','Width',Width);
        ini.WriteInteger('Panel','Height',Height);
      end;
    finally
      ini.Free;
    end;
  except
  end;
end;


procedure TFormLua.ActionRunExecute(Sender: TObject);
begin
  Memo1.Lines.Add(rsStartScript);
  LuaReInstance.Enabled:=False;
  ActionRun.Enabled:=False;
  SynEdit1.ReadOnly:=True;
  try
  lastkey:=0;
  TLuaHttpExpr(Worker).do_string(SynEdit1.Lines.Text);
  finally
    SynEdit1.ReadOnly:=False;
    ActionRun.Enabled:=True;
    LuaReInstance.Enabled:=True;
  end;
  Memo1.Lines.Add(rsEndScript);
end;

procedure TFormLua.ActionTempHttpExecute(Sender: TObject);
begin
  with CreateMessageDialog(rsLostYourScri, mtWarning, mbYesNo) do begin
    try
      Position:=poOwnerFormCenter;
      if ShowModal=mrYes then
        SynEdit1.Lines.Text:=''; //StrHolder1.Strings.Text;
    finally
      Free;
    end;
  end;
end;

procedure TFormLua.EditCodeComp1Execute(Sender: TObject);
begin
  SynEdit1.CommandProcessor(SynCompletion1.ExecCommandID,'',nil);
end;


procedure TFormLua.EditCopy1Execute(Sender: TObject);
begin
  SynEdit1.CopyToClipboard;
end;

procedure TFormLua.EditCut1Execute(Sender: TObject);
begin
  SynEdit1.CutToClipboard;
end;

procedure TFormLua.EditDelete1Execute(Sender: TObject);
begin
  SynEdit1.ClearSelection;
end;

procedure TFormLua.EditInsertPathExecute(Sender: TObject);
var
  NewPath:string;
  i,len:Integer;
  ch:char;
begin
  if EditPathDlg.Execute then begin
    i:=1;
    len:=Length(EditPathDlg.FileName);
    NewPath:='';
    while i<=len do begin
      ch:=EditPathDlg.FileName[i];
      NewPath:=NewPath+ch;
      if ch='\' then
        NewPath:=NewPath+ch;
      Inc(i);
    end;
    SynEdit1.InsertTextAtCaret(NewPath);
  end;
end;

procedure TFormLua.EditPaste1Execute(Sender: TObject);
begin
  SynEdit1.PasteFromClipboard;
end;

procedure TFormLua.EditPaste1Update(Sender: TObject);
begin
  TAction(Sender).Enabled:=SynEdit1.CanPaste;
end;

procedure TFormLua.EditRedo1Execute(Sender: TObject);
begin
  SynEdit1.Redo;
end;

procedure TFormLua.EditRedo1Update(Sender: TObject);
begin
  TAction(Sender).Enabled:=SynEdit1.CanRedo;
end;

procedure TFormLua.EditSelectAll1Execute(Sender: TObject);
begin
  SynEdit1.SelectAll;
end;

procedure TFormLua.EditUndo1Execute(Sender: TObject);
begin
  SynEdit1.Undo;
end;

procedure TFormLua.EditUndo1Update(Sender: TObject);
begin
  TAction(Sender).Enabled:=SynEdit1.CanUndo;
end;

procedure TFormLua.FileNew1Execute(Sender: TObject);
begin
  if CheckModified then begin
    SynEdit1.ClearAll;
    SynEdit1.Modified:=False;
    FileOpenDlg.FileName:='';
    SetCaption('');
  end;
end;

procedure TFormLua.FileOpen1Execute(Sender: TObject);
begin
  if CheckModified then begin
    if FileOpenDlg.Execute then begin
      SynEdit1.Lines.LoadFromFile(FileOpenDlg.FileName);
      SynEdit1.Modified:=False;
      FileSaveAs1.Dialog.FileName:=FileOpenDlg.FileName;
      SetCaption(FileOpenDlg.FileName);
    end;
  end;
end;

procedure TFormLua.FileSaveAs1Accept(Sender: TObject);
begin
  SynEdit1.Lines.SaveToFile(FileSaveAs1.Dialog.FileName);
  FileOpenDlg.FileName:=FileSaveAs1.Dialog.FileName;
  SetCaption(FileOpenDlg.FileName);
  SynEdit1.Modified:=False;
end;

procedure TFormLua.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  CanClose:=CheckModified;
end;

procedure TFormLua.ActionCopyExecute(Sender: TObject);
begin
  WorkGrid.CopyToClipboard(True);
end;

procedure TFormLua.ActionExportExecute(Sender: TObject);
var
  IsUTF8:Boolean;
  delimit:char;
begin
  if SaveDialog1.Execute then begin
    if UpperCase(ExtractFileExt(SaveDialog1.FileName))='.CSV' then begin
      //WorkGrid.SaveToCSVFile(SaveDialog1.FileName,',')
      if Assigned(Worker) then begin
        IsUTF8:=TLuaHttpExpr(Worker).FuseUTF8String;
        delimit:=TLuaHttpExpr(Worker).CSVDelimit;
        end else begin
          IsUTF8:=True;
          delimit:=',';
        end;
      GridSaveCSVFile(SaveDialog1.FileName,delimit,IsUTF8)
      end else
        WorkGrid.SaveToFile(SaveDialog1.FileName);
  end;
end;

procedure TFormLua.ActionImportExecute(Sender: TObject);
var
  IsUTF8:Boolean;
  delimit:char;
begin
  if OpenDialog1.Execute then begin
    if UpperCase(ExtractFileExt(OpenDialog1.FileName))='.CSV' then begin
      //WorkGrid.LoadFromCSVFile(OpenDialog1.FileName,',')
      if assigned(Worker) then begin
         IsUTF8:=TLuaHttpExpr(Worker).FuseUTF8String;
         delimit:=TLuaHttpExpr(Worker).CSVDelimit;
         TLuaHttpExpr(Worker).StrGrid_clear;
         end else begin
           IsUTF8:=True;
           delimit:=',';
           WorkGrid.ColCount:=1;
           WorkGrid.RowCount:=1;
         end;
      GridLoadCSVFile(OpenDialog1.FileName,delimit,IsUTF8);
      end else
        WorkGrid.LoadFromFile(OpenDialog1.FileName);
  end;
end;

procedure TFormLua.ActionInsRegexExecute(Sender: TObject);
begin
  SynEdit1.InsertTextAtCaret(Format(rsVregexRegEx_, [LineEnding, LineEnding, LineEnding, LineEnding]));
end;

procedure TFormLua.ActionInsRegexMatchExecute(Sender: TObject);
begin
  SynEdit1.InsertTextAtCaret(Format(rsVregexRegEx2_, [LineEnding, LineEnding, LineEnding, LineEnding]));
end;

procedure TFormLua.FormDestroy(Sender: TObject);
begin
  Worker.Free;
  SavePosition;
end;

procedure TFormLua.FormDropFiles(Sender: TObject;
  const FileNames: array of String);
begin
  if CheckModified then begin
    SynEdit1.Lines.LoadFromFile(FileNames[0]);
    SynEdit1.Modified:=False;
  end;
end;

procedure TFormLua.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if not LuaReInstance.Enabled then
    lastkey:=key;
end;

procedure TFormLua.FormShow(Sender: TObject);
var
  i:Integer;
  bRun:Boolean;
begin
  SynEdit1.Lines.Clear;

  WorkGrid.FixedRows:=1;
  WorkGrid.FixedCols:=1;

  if Paramcount>0 then begin
    FileOpenDlg.FileName:=ParamStr(1);
    SynEdit1.Lines.LoadFromFile(ParamStr(1));
    bRun:=False;
    for i:=2 to Paramcount do
      if UpperCase(ParamStr(i))='-R' then
        bRun:=True;
    if bRun then
      ActionRun.Execute;
  end;
end;

procedure TFormLua.LuaReInstanceExecute(Sender: TObject);
begin
  Worker.Free;
  Sleep(100);
  Worker:=TLuaHttpExpr.Create(True);
  Memo1.Lines.Add(rsLuaInstanceR);
end;

procedure TFormLua.SynCompletion1CodeCompletion(var Value: string;
  SourceValue: string; var SourceStart, SourceEnd: TPoint; KeyChar: TUTF8Char;
  Shift: TShiftState);
var
  po:TPoint;
  s,v:string;
  i,l:Integer;
begin
  po:=SourceStart;
  Dec(po.x);
  s:=SynCompletion1.Editor.GetWordAtRowCol(po);
  if s<>'' then begin
    s:=UpperCase(s+'.');
    v:=UpperCase(Value);
    l:=Length(s);
    i:=Pos(s,v);
    if i>0 then
      Delete(Value,i,l);
  end;
end;

procedure TFormLua.SetCaption(str: string);
begin
  Caption:=Format(rsLuaGridS, [ExtractFileName(str)]);
end;

procedure TFormLua.WorkGridCompareCells(Sender: TObject; ACol, ARow, BCol,
  BRow: Integer; var Result: integer);
var
  A,B:Integer;
begin
  if ACol=0 then begin
    if ARow=0 then
      A:=-1
      else
        A:=StrToIntDef(WorkGrid.Cells[ACol,ARow],0);
    if BRow=0 then
      B:=-1
      else
        B:=StrToIntDef(WorkGrid.Cells[BCol,BRow],0);
    if WorkGrid.SortOrder=soAscending then
      Result:=A-B
      else
        Result:=B-A;
    end else begin
      if WorkGrid.SortOrder=soAscending then
        Result:=CompareStr(WorkGrid.Cells[ACol,ARow],WorkGrid.Cells[BCol,BRow])
        else
          Result:=CompareStr(WorkGrid.Cells[BCol,BRow],WorkGrid.Cells[ACol,ARow]);
    end;
end;


function TFormLua.CheckModified:boolean;
var
  mr:Integer;
begin
  Result:=True;
  if SynEdit1.Modified then begin
    with CreateMessageDialog(rsScriptIsModi, mtWarning, mbYesNoCancel) do begin
      try
        Position:=poOwnerFormCenter;
        mr:=ShowModal;
        if mr=mrYes then begin
          if FileOpenDlg.FileName<>'' then
            FileSaveAs1.Dialog.FileName:=FileOpenDlg.FileName;
          FileSaveAs1.Execute;
          Result:=FileSaveAs1.Dialog.UserChoice<>mrCancel;
          if FileSaveAs1.Dialog.UserChoice=mrOK then begin
            FileOpenDlg.FileName:=FileSaveAs1.Dialog.FileName;
            SetCaption(FileOpenDlg.FileName);
          end;
        end else
          Result:=mr=mrNo;
      finally
        Free;
      end;
    end;
  end;
end;


end.

