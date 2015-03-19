ff=Find_FileFirst('e:\\temp\\')
j=0
Grid_Clear()
if ff~=nil then
repeat
name=Find_FileName(ff)
Grid_Value(0,j,name)
j=j+1
until not Find_FileNext(ff)
end
Find_FileClose(ff)
