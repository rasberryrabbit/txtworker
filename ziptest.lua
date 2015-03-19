if true==Zip_AddText('test.zip','','_text_','it is test') then
write('add ok')
end
if 'it is test'==Zip_ExtractText('test.zip','','_text_') then
write('extract ok')
end
if true==Zip_Add('testfile.zip','','*.txt') then
write('add file ok')
end;
