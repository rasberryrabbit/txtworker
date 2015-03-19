ff=Find_FolderFirst('d:\\*.*')
if ff~=nil then
ii=0
repeat
name=Find_FolderName(ff);
write(name)
ii=ii+1
if ii>500 then
break
end
until not Find_FolderNext(ff)
end
Find_FolderClose(ff)
