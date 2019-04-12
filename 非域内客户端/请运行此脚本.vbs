Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set MyFile = objfso.CreateTextFile("wlan.bat", True)
Set shell=createobject("wscript.shell")

Shell.run ("create.bat")
wscript.sleep 5000

Set objTextFile = objFSO.OpenTextFile("interface.txt", ForReading)

For i = 1 to 2
objTextFile.ReadLine
Next

strLine = objTextFile.ReadLine
strMid = mid(strLine,8)

MyFile.WriteLine("wlan.exe sp" & " " & strMid & " " & "MideaNoDomain.xml")
MyFile.Close
objTextFile.Close

Shell.run "wlan.bat"
wscript.sleep 5000
set shell=nothing
Set DFile = objfso.GetFile("wlan.bat")
Dfile.Delete
Set DFile2 = objfso.GetFile("interface.txt")
Dfile2.Delete
Wscript.echo "无线WIFI已完成设置!"