Dim name
' name = "N00_000_Hair_00_HAIR_01 (Instance).png"
' name = "N00_001_02_Bottoms_01_CLOTH (Instance).png"
name = "N00_002_03_Tops_01_CLOTH_01 (Instance).png"

function convertHair(name)
    Dim str
    str = mid(name, 21, 3) ' _01
    convertHair = "H" & str
end function

' function convertCloth(name)
'     Dim str
'     str = mid(name, 6, 5) ' 01_01
'     convertCloth = "C_" & str
' end function

' C_010101
function convertCloth(name)
    Dim str, noNumberCloth
    noNumberCloth = 0
    ' N00_002_03_Tops_01_CLOTH_01 (Instance).png
    ' str = Mid(name, 6, 5) ' 01_01
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

' msgbox convertHair(name)
msgbox convertCloth(name)
' msgbox convertUnknown(23)
' msgbox addX("BDSKN", 11)