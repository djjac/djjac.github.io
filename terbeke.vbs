input= InputBox("Voer aatal bakken in")





If IsNumeric(input) Then

pallet = input / 1440
if pallet  > 0 then
pallet = Int(pallet)
end if

if pallet > 0 then plus =" +" else plus =" " 

doos =  ( input - (pallet * 1440))/ 6
doos = Int (doos) 
bakken =  input - (pallet * 1440 ) - doos * 6



if pallet > 0 then 
pallet = pallet & " pallet"
else 
pallet = ""
end if



if doos > 0 then 
if doos > 1 then 
hoog = doos/16
hoog = Int (hoog) 
resthoog = doos - (hoog*16)

hoog = hoog & " hoog"
if resthoog > 0 then
resthoog = " + " & resthoog & " dozen" 
end if

doos = plus & doos &  " dozen ( " & hoog &  resthoog  &  " )"
else 
doos = plus & doos &  " doos" 
end if
else doos = ""
end if






if bakken > 0 then 
if bakken > 1 then 
bakken = " + " & bakken & " bakken  " 
else 
bakken = " + " & bakken & " bak  " 
end if
else bakken = ""
end if


output = pallet & doos & bakken 
 

inputdec= input/6
inputdec= FormatNumber(inputdec,2)

if msgBox(output, vbretrycancel, inputdec & " dozen") = vbretry then
Set WshShell = WScript.CreateObject ("WScript.Shell")
WshShell.Run ("terbeke.vbs")


end if


Else
   Wscript.Echo "waardes moeten numeriek zijn" 
Set WshShell = WScript.CreateObject ("WScript.Shell")
WshShell.Run ("terbeke.vbs")


End If