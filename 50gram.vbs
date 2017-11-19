input= InputBox("Voer aatal kratten in","50 gram flakes")
input2= InputBox("Vooraad","50 gram flakes")





If IsNumeric(input) and IsNumeric(input2)Then

 vrd = input2
   ank = input * 12
   totaal = ank - vrd
   kilos = totaal * 0.1
   t_input = Round((kilos)/15,1)

output1 = input & " kratten - " & vrd & " vooraad  = " & totaal & " bakken "
output = output1 & vbNewLine & vbNewLine & kilos & " kg = " & t_input & " kratten"
 

inputdec= input2


if msgBox(output, vbretrycancel, "50 gram AH") = vbretry then
Set WshShell = WScript.CreateObject ("WScript.Shell")
WshShell.Run ("50gram.vbs")
end if


Else
   Wscript.Echo "waardes moeten numeriek zijn" 
Set WshShell = WScript.CreateObject ("WScript.Shell")
WshShell.Run ("50gram.vbs")


End If