Grid_Clear();
Clear();
txt=GetHttpText('http://www.daum.net');
if not txt then
write('error1')
end;
reg=RegEx_New('(?is)<li\\s+class=\"rank_li[^\"]+?\"><div\\s+class=\"[^\"]+\"><a\\s+href=\"([^\"]+)\"(\\s+title=\"([^\"]+)\")?\\s+class=\"@\\d+\"([^>]+)?>(<strong>)?([^>]+)(</strong>)?</a><span\\s+class=\"[^\"]+\">([^>]+)</span>(<span\\s+class=\"[^\"]+\">([^>]+)</span>|)?</div></li>');
--str,pos,len=RegEx_Match(reg,txt);
--print(str,pos,len);
res=RegEx_MatchAll(reg,txt);
xmax=res['Col'];
ymax=res['Row'];
for i=0,ymax do
for j=0,xmax do
  key=res[tostring(i)..','..tostring(j)];
  if key then
    Grid_Value(j,i,key);
  else
    break
  end;
end;
end;
RegEx_Delete(reg); 
