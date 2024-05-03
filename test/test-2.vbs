Dim name
' name = "N00_000_Hair_00_HAIR_01 (Instance).png"
name = "N00_001_02_Bottoms_01_CLOTH (Instance).png"

function convertHair(name)
    Dim str
    str = mid(name, 21, 3) ' _01
    convertHair = "H" & str
end function

function convertCloth(name)
    Dim str
    str = mid(name, 6, 5) ' 01_01
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

' msgbox convertHair(name)
' msgbox convertCloth(name)
msgbox convertUnknown(23)