Dim objFSO
Dim objFile
Dim objTextFile
Dim strText
Dim strPath
Dim strProjectPath
Dim unmatchCount
Dim characters
unmatchCount = 0
characters = 11 ' the length of the file name

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
strPath = objFSO.getParentFolderName(WScript.ScriptFullName) & "\files"
strProjectPath = objFSO.getParentFolderName(WScript.ScriptFullName) & "\result.txt"
Set objTextFile = objFSO.CreateTextFile(strProjectPath)

' H_01
function convertHair(name)
    Dim str
    str = Mid(name, 21, 3) ' _01
    convertHair = "H" & str
end function

' C_010101
function convertCloth(name)
    Dim str, noNumberCloth
    noNumberCloth = 0
    ' N00_002_03_Tops_01_CLOTH_01 (Instance).png
    str = Mid(name, 6, 2) & Mid(name, 9, 2) ' 0101
    if InStr(1, name, "CLOTH (", vbTextCompare) > 0 Then
        if noNumberCloth < 10 Then
            str = str & "0" & noNumberCloth
        else
            str = str & noNumberCloth ' 010100
        end if
        noNumberCloth = noNumberCloth + 1
    else
        str = str & Mid(name, Len(name) - 16, 2) ' 010101
    end if
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
        convertUnknown = str & Left(str, 3)
    end if
end function

function addX(name, max)
    Dim length, x, xs, i
    length = Len(name)
    x = max - length

    if x < 1 Then
        addX = ""
    else
        for i = 2 to x
            xs = xs + "X"
        next
    addX = "_" & xs
    end if
end function

function convertName(name)
    Dim str

    if Len(name) < 10 Then
        convertName = name
        Exit function
    elseif InStr(1, name, "copy", vbTextCompare) > 0 Then
        convertName = name
        Exit function
    elseif InStr(1, name, "Body_00_SKIN", vbTextCompare) > 0 Then
        str = "BDSKN"
    elseif InStr(1, name, "EyeHighlight", vbTextCompare) > 0 Then
        str = "EHL"
    elseif InStr(1, name, "EyeIris", vbTextCompare) > 0 Then
        str = "EI"
    elseif InStr(1, name, "EyeWhite", vbTextCompare) > 0 Then
        str = "EW"
    elseif InStr(1, name, "FaceBrow", vbTextCompare) > 0 Then
        str = "FB"
    elseif InStr(1, name, "FaceEyelash", vbTextCompare) > 0 Then
        str = "FELSH"
    elseif InStr(1, name, "FaceEyeline", vbTextCompare) > 0 Then
        str = "FEL"
    elseif InStr(1, name, "FaceMouth", vbTextCompare) > 0 Then
        str = "FM"
    elseif InStr(1, name, "Face_00_SKIN", vbTextCompare) > 0 Then
        str = "FS"
    elseif InStr(1, name, "HairBack", vbTextCompare) > 0 Then
        str = "HB"
    elseif InStr(1, name, "HAIR_", vbTextCompare) > 0 Then
        str = convertHair(name)
    elseif InStr(1, name, "CLOTH", vbTextCompare) > 0 Then
        str = convertCloth(name)
    else
        convertName = name
        Exit function
    end if

    str = str & addX(str, characters) ' BDSKN_XXXXX
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
    Dim converted, objFile2
    converted = convertName(objFile.Name)
    
    objTextFile.WriteLine(objFile.Name & " => " & converted)

    If objFile.Name <> converted Then
        objFile.Name = converted
    End If   
Next

Set objFSO = Nothing