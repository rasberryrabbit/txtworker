program lua_text_test;

uses sysutils,lua,lua52,uLuaHttpExpr;

var
  L:TLua;
begin
  L:=TLuaHttpExpr.Create;
  try
    L.DoFile('test.lua');
  finally
    L.Free;
  end;
  readln;
end.

