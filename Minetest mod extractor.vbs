'V1

set sa = CreateObject("Shell.Application")
set sf=CreateObject("Scripting.FileSystemObject")
pp=sf.GetAbsolutePathName("")

if not sf.FolderExists("extracted") then
	sf.CreateFolder("extracted")
end if

for each fi1 in sf.GetFolder(pp).Files
	if instr(fi1.name,".zip")<>0 then
		sa.NameSpace(pp).CopyHere sa.NameSpace(pp & "\" & fi1.name).items
		sf.movefile fi1.name, "extracted\" + fi1.name
	end if
next

for each fo1 in sf.GetFolder(pp).subfolders
	if instr(fo1.name,"-master")<>0 then
		sf.moveFolder fo1, replace(fo1.name,"-master","")
	end if
next