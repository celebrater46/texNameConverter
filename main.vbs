' https://step-learn.com/article/vbscript/059-file-nama-change.html

' Dim strFilePath

' ' type "J:\Dropbox\PC5_cloud\pg\VB\testVBS\test\test.txt"
' strFilePath = inputbox("type the target file directory (includes the file name).", "INPUT BOX")
' Set objFS = CreateObject("Scripting.FileSystemObject")
' Set objFile = objFS.GetFile(strFilePath)

' ' changing the file name
' objFile.Name = "test-dagya.txt"





Dim objFSO, objFile, objTextFile, strText, strPath, strProjectPath, unmatchCount
unmatchCount = 0



Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
' strPath = inputbox("Input the target directory.", "INPUT BOX")
strPath = objFSO.getParentFolderName(WScript.ScriptFullName) & "\files"
strProjectPath = objFSO.getParentFolderName(WScript.ScriptFullName) & "\files.txt"
Set objTextFile = objFSO.CreateTextFile(strProjectPath)

' H_01
function convertHair(name)
    Dim str
    str = mid(name, 21, 3) ' _01
    convertHair = "H" & str
end function

' C_01_01
function convertCloth(name)
    Dim str
    str = mid(name, 6, 5) ' 01_01
    convertCloth = "C_" & str
end function

' When not found from the file patterns
function convertUnknown(num)
    Dim str
    str = "UNKNOWN_"

    if num < 10 Then
        convertUnknown = str & "00" & num
    elseif num < 100 Then
        convertUnknown = str & "0" & num
    elseif num < 1000 Then
        convertUnknown = str & num
    else
        convertUnknown = str & left(str, 3)
    end if
end function

function convertName(name)
    Dim str
    if InStr(name, "Body_00_SKIN") > 0 Then
        str = "BDSKN"
    elseif InStr(name, "EyeHighlight") > 0 Then
        str = "EHL"
    elseif InStr(name, "EyeIris") > 0 Then
        str = "EI"
    elseif InStr(name, "EyeWhite") > 0 Then
        str = "EW"
    elseif InStr(name, "FaceBrow") > 0 Then
        str = "FB"
    elseif InStr(name, "FaceEyelash") > 0 Then
        str = "FELSH"
    elseif InStr(name, "FaceEyeline") > 0 Then
        str = "FEL"
    elseif InStr(name, "FaceMouth") > 0 Then
        str = "FM"
    elseif InStr(name, "Face_00_SKIN") > 0 Then
        str = "FS"
    elseif InStr(name, "HairBack") > 0 Then
        str = "HB"
    elseif InStr(name, "HAIR_") > 0 Then
        str = convertHair(name)
    elseif InStr(name, "CLOTH") > 0 Then
        str = convertCloth(name)
    else
        ' str = convertUnknown(unmatchCount)
        ' unmatchCount = unmatchCount + 1
        str = Replace(name, ".png", "")
    end if

    convertName = str & ".png"

    ' N00_000_00_Body_00_SKIN (Instance).png			->	BDSKN.png
    ' N00_000_00_EyeHighlight_00_EYE (Instance).png	->	EHL.png
    ' N00_000_00_EyeIris_00_EYE (Instance).png		->	EI.png
    ' N00_000_00_EyeWhite_00_EYE (Instance).png		->	EW.png
    ' N00_000_00_FaceBrow_00_FACE (Instance).png	->	FB.png
    ' N00_000_00_FaceEyelash_00_FACE (Instance).png	->	FELSH.png
    ' N00_000_00_FaceEyeline_00_FACE (Instance).png	->	FEL.png
    ' N00_000_00_FaceMouth_00_FACE (Instance).png	->	FM.png
    ' N00_000_00_Face_00_SKIN (Instance).png			->	FS.png
    ' N00_000_00_HairBack_00_HAIR (Instance).png		->	HB.png
    ' N00_000_Hair_00_HAIR_01 (Instance).png			->	H_01.png
    ' N00_000_Hair_00_HAIR_03 (Instance).png			->	H_03.png
    ' N00_000_Hair_00_HAIR_04 (Instance).png			->	H_04.png
    ' N00_000_Hair_00_HAIR_05 (Instance).png			->	H_05.png
    ' N00_000_Hair_00_HAIR_06 (Instance).png			->	H_06.png
    ' N00_001_02_Bottoms_01_CLOTH (Instance).png	->	C_01.png
    ' N00_002_03_Tops_01_CLOTH_01 (Instance).png	->	C_02_01.png
    ' N00_002_03_Tops_01_CLOTH_02 (Instance).png	->	C_02_02.png
    ' N00_002_03_Tops_01_CLOTH_03 (Instance).png	->	C_02_03.png
    ' N00_007_01_Tops_01_CLOTH_01 (Instance).png	->	C_07_01.png
    ' N00_007_01_Tops_01_CLOTH_02 (Instance).png	->	C_07_02.png
    ' N00_008_01_Shoes_01_CLOTH (Instance).png		->	C_08_01.png
end function

For Each objFile In objFSO.GetFolder(strPath).Files
    ' File Name
    If strText <> "" Then ' is not first?
        ' strText = strText & vbCrLf & objFile.Name
        strText = strText & vbCrLf & convertName(objFile.Name)
    Else
        ' strText = objFile.Name ' is first
        strText = convertName(objFile.Name) ' is first
    End If
    
Next

objTextFile.WriteLine(strText)

Set objFSO = Nothing