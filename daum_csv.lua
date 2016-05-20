Grid_Clear();
txt=GetHttpText('http://www.daum.net');
if not txt then
write('error')
end;
reg=RegEx_New('(?i)<span(\\s+)?class=\"txt_issue\">(\\s+)?<a href=\"[^\"]+\" class=\"[^\"]+\">\\s+?(.+)\\s+?</a>\\s+</span>\\s+<em class[^>]+>(.+)\\S+');
res=RegEx_MatchAllCSV(reg,txt);
Grid_LoadStr(res)
RegEx_Delete(reg);  
