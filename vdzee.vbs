
datum = DATE()


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")

Set aantal = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.ACE.OLEDB.12.0; " & _
        "Data Source = C:\Users\djjac\Documents\streaky.mdb" 

aantal.Open "SELECT SUM([Aantal verpakkingen]) as totaal FROM Bestellingen WHERE Klantnummer = 17531 AND Verzenddatum BETWEEN DATE() AND DateAdd(""d"",8,Date())  And Artikelnummer =40263",_
     objConnection, adOpenStatic, adLockOptimistic

aantal.MoveFirst

Do Until aantal.EOF
  aantal1= aantal.Fields(0).value
   aantal.MoveNext
Loop


Set aantal2 = CreateObject("ADODB.Recordset")



aantal2.Open "SELECT SUM([Aantal verpakkingen]) as nog FROM Bestellingen WHERE [Gereed?]= No AND Klantnummer = 17531 AND Verzenddatum BETWEEN DATE() AND DateAdd(""d"",8,Date()) And Artikelnummer =40263",_
     objConnection, adOpenStatic, adLockOptimistic

aantal2.MoveFirst

Do Until aantal2.EOF
  aantal2ok= aantal2.Fields(0).value
   aantal2.MoveNext
Loop
aantal2.Close
aantal.Close

af= aantal1 - aantal2ok

 Wscript.Echo    datum & vbNewLine& vbNewLine& af &"/"&aantal1&" bakken af. "&vbNewLine& "Nog "& aantal2ok &" bakken. " 


objConnection.Close	