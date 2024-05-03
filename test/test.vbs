dim name, str
name = inputbox("type some file name...")

' com is par
' select case name
'     case InStr(name, "Body_00_SKIN") > 0
'         str = "BDSKN"
'     case InStr(name, "EyeHighlight") > 0
'         str = "EHL"
'     case InStr(name, "EyeIris") > 0
'         str = "EI"
'     case else 
'         str = "_UNKNOWN_"
' end select 

if InStr(name, "Body_00_SKIN") > 0 Then
    str = "BDSKN"
elseif InStr(name, "EyeHighlight") > 0 Then
    str = "EHL"
elseif InStr(name, "EyeIris") > 0 Then
    str = "EI"
else
    str = "_UNKNOWN_"
end if

msgbox str