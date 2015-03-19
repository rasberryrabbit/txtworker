Grid_Clear();
txt=GetHttpText('http://www.daum.net');
if not txt then
write('error')
end;
reg=RegEx_New('(?i)<span(\\s+)?class=\"txt_issue\">(\\s+)?<a href=\"[^\"]+\" class=\"[^\"]+\">(\\s+)?(\\S+)');
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
