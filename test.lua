txt=GetHttpText('http://www.daum.net');
if not txt then
print('error')
end;
reg=RegEx_New('(?is)<li\\s+class=\"rank_li[^\"]+?\"><div\\s+class=\"[^\"]+\"><a\\s+href=\"([^\"]+)\"(\\s+title=\"([^\"]+)\"|)?\\s+class=\"@\\d+\">(<strong>|)?([^>]+)(</strong>|)?</a><span\\s+class=\"[^\"]+\">([^>]+)</span>(<span\\s+class=\"[^\"]+\">([^>]+)</span>|)?</div></li>');
--str,pos,len=RegEx_Match(reg,txt);
--print(UTF8Decode(str),pos,len);
res=RegEx_MatchAll(reg,txt);
for i=0,50 do
j=0;
  key=res[tostring(i)..','..tostring(j)];
  if key then
    print(tostring(i)..','..tostring(j)..'='..UTF8Decode(key));
  else
    break
  end;
end;
RegEx_Delete(reg);
