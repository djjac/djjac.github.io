input= InputBox("Voer aatal bakken in")





If IsNumeric(input) Then

pallet = input / 240
if pallet  > 0 then
pallet = Int(pallet)
end if

if pallet > 0 then plus =" +" else plus =" " 

krat =  ( input - (pallet * 240))/ 6
krat = Int (krat) 
bakken =  input - (pallet * 240 ) - krat * 6



if pallet > 0 then 
pallet = pallet & " pallet"
else 
pallet = ""
end if



if krat > 0 then 
if krat > 1 then 
krat = plus & krat &  " kratten" 
else 
krat = plus & krat &  " krat" 
end if
else krat = ""
end if

if bakken > 0 then 
if bakken > 1 then 
bakken = " + " & bakken & " bakken  " 
else 
bakken = " + " & bakken & " bak  " 
end if
else bakken = ""
end if


output = pallet & krat & bakken 
Msgbox output, 1, ""




Else
    Wscript.Echo "Beide waardes moeten numeriek zijn" 

End If