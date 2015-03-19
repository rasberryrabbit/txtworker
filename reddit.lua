url='http://www.reddit.com';
-- start loop
txt = GetHttpText(url);
if txt==nil then
error('error');
end;
-- get next url
reg=RegEx_New('(?is)<a\\s+href=\"([^\"]+)\"\\s+rel=\"nofollow next\" >');
res=RegEx_MatchAll(reg,txt);
url=res['0,1'];
Grid_Clear();
regb=RegEx_New('<p\\s+class=\"title\"><a\\s+class=\"([^\"]+)\"\\s+href=\"([^\"]+)"\\s+tabindex=\"\\d+\"\\s+>([^>]+)</a>[^>]+<span class=\"domain\">');
body=RegEx_MatchAll(regb,txt);
Grid_Clear();
--Grid_FromTable(body);
for i=0,15 do
j=0;
while j<300 do
  value=body[ tostring(j)..','.. tostring(i) ];
  if value then
    Grid_Value(1,j,value)
  else
    break
  end;
  j=j+1;
end
end
regc=RegEx_New('<a\\s+href=\"[^\"]+\"\\s+class=\"[^\"]+\"\\s+>([^>]+)</a><span\\s+class=\"(^\"]+)\"></span>[^>]+<a\\s+href=\"[^\"]+\"\\s+class=\"[^\"]+\"\\s+>[^>]+</a>');
body2=RegEx_MatchAll(regc,txt);
--Grid_FromTable(body2)
for i=0,15 do
j=0;
while j<300 do
  value=body2[ tostring(j)..','.. tostring(i) ];
  if value then
    Grid_Value(1,j,value)
  else
    break
  end;
  j=j+1;
end
end
RegEx_Delete(regc)
RegEx_Delete(regb);
RegEx_Delete(reg);



