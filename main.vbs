' https://step-learn.com/article/vbscript/059-file-nama-change.html

Dim strFilePath

' type "J:\Dropbox\PC5_cloud\pg\VB\testVBS\test\test.txt"
strFilePath = inputbox("type the target file directory (includes the file name).", "INPUT BOX")
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.GetFile(strFilePath)

' changing the file name
objFile.Name = "test-dagya.txt"