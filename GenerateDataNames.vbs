Option Explicit
Dim objexcelR, objexcelW, objWorkbookR, objWorkbookW, objDriverSheetR, objDriverSheetW, objDicGen, objDicCity
Dim lstAlfa
Dim numNames, numApell
Dim Iterator
Dim apPat, apMat, fechaNac, intGenero, intCiudad, genero, nombre, ciudad, CURP, RFC

set objexcelR = Createobject("Excel.Application")
set objexcelW = Createobject("Excel.Application")
set objWorkbookR = objExcelR.WorkBooks.Open(getPath & "\nombresApellidos.xlsx")
Set objWorkbookW = objExcelW.WorkBooks.Add(getPath & "\nombresApellidos_GEN.xlsx")
Set objDriverSheetR = objWorkbookR.Worksheets("Names")
Set objDriverSheetW = objWorkbookW.Worksheets("Names")
Set objDicGen = createobject("Scripting.Dictionary")
Set objDicCity = createobject("Scripting.Dictionary")

objDicGen.Add "1", "H"
objDicGen.Add "2", "M"

objDicCity.Add "1", "AS"
objDicCity.Add "2", "BC"
objDicCity.Add "3", "BS"
objDicCity.Add "4", "CC"
objDicCity.Add "5", "CS"
objDicCity.Add "6", "CH"
objDicCity.Add "7", "CL"
objDicCity.Add "8", "CM"
objDicCity.Add "9", "DF"
objDicCity.Add "10", "DG"
objDicCity.Add "11", "GT"
objDicCity.Add "12", "GR"
objDicCity.Add "13", "HG"
objDicCity.Add "14", "JC"
objDicCity.Add "15", "MC"
objDicCity.Add "16", "MN"
objDicCity.Add "17", "MS"
objDicCity.Add "18", "NT"
objDicCity.Add "19", "NL"
objDicCity.Add "20", "OC"
objDicCity.Add "21", "PL"
objDicCity.Add "22", "QO"
objDicCity.Add "23", "QR"
objDicCity.Add "24", "SP"
objDicCity.Add "25", "SL"
objDicCity.Add "26", "SR"
objDicCity.Add "27", "TC"
objDicCity.Add "28", "TS"
objDicCity.Add "29", "TL"
objDicCity.Add "30", "VZ"
objDicCity.Add "31", "YN"
objDicCity.Add "32", "ZS"

lstAlfa = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

numNames = 100
numApell = 441

objexcelW.Visible = True
objexcelR.Visible = True

For Iterator = 1 To 501 Step 1

	apPat = trim(objDriverSheetR.cells(RandomNumber(2,numApell), 3))
	apMat = trim(objDriverSheetR.cells(RandomNumber(2,numApell), 3))
	fechaNac = Cstr(RandomNumber(1950,1999)) & "/" & Left("0" + Cstr(RandomNumber(1,12)), 2) & "/" & Left("0" + Cstr(RandomNumber(1,28)), 2)
	intGenero = RandomNumber(1,2)
	intCiudad = RandomNumber(1,32)
	genero = objDicGen.Item(Cstr(intGenero))
	nombre = trim(objDriverSheetR.cells(RandomNumber(2,numNames) , intGenero))
	ciudad = objDicCity.Item(Cstr(intCiudad))
	CURP = getCURP(nombre, apPat, apMat, fechaNac, genero, ciudad)
	RFC = getRFC(nombre, apPat, apMat, fechaNac)
	objDriverSheetW.cells(Iterator+1, 1).Value = nombre
	objDriverSheetW.cells(Iterator+1, 2).Value = apPat
	objDriverSheetW.cells(Iterator+1, 3).Value = apMat
	objDriverSheetW.cells(Iterator+1, 4).Value = fechaNac
	objDriverSheetW.cells(Iterator+1, 5).Value = objDriverSheetR.cells(intCiudad+1, 4)
	objDriverSheetW.cells(Iterator+1, 6).Value = genero
	objDriverSheetW.cells(Iterator+1, 7).Value = CURP
	objDriverSheetW.cells(Iterator+1, 8).Value = RFC
Next
objexcelW.ActiveWorkbook.SaveAs getPath & "\nombresApellidos_GEN1.xlsx"
objexcelW.ActiveWorkbook.close
objexcelR.ActiveWorkbook.close

Function RandomNumber(nLow, nHigh)
	WScript.Sleep 003
	Set objRandom = CreateObject("System.Random")
	RandomNumber = objRandom.Next_2(nLow, nHigh)
End Function

'**********************************************************
'Name function: getCURP
'Parameters:
'	firtsName
'	lastName
'	OtherName
'	birthDate: AAAA/MM/DD
'**********************************************************
Public Function getCURP(firtsName, lastName, OtherName, birthDate, gender, city)
	Dim preCURP
	preCURP = Mid(lastName, 1, 2) & Mid(OtherName, 1, 1) & Mid(firtsName, 1, 1) & Mid(birthDate, 3, 2) & Mid(birthDate, 6, 2) & Mid(birthDate, 9, 2) & gender & city & cons(lastName) & cons(OtherName) & cons(firtsName) & "0" & RandomNumber(1, 9)
	getCURP = Ucase(preCURP)
End Function

Public Function getRFC(firtsName, lastName, OtherName, birthDate)
	Dim preRFC
	preRFC = Mid(lastName, 1, 2) & Mid(OtherName, 1, 1) & Mid(firtsName, 1, 1) & Mid(birthDate, 3, 2) & Mid(birthDate, 6, 2) & Mid(birthDate, 9, 2) & Mid(lstAlfa, RandomNumber(1,36), 1) & Mid(lstAlfa, RandomNumber(1,36), 1) & Mid(lstAlfa, RandomNumber(1,36), 1)
	getRFC = Ucase(preRFC)
End Function

Private Function cons(strString)
	Dim Iterator, isvowel
	
	For Iterator = 2 To len(strString) Step 1
		isvowel = Ucase(Mid(strString, Iterator, 1))
		If (isvowel <> "A" and isvowel <> "E" and isvowel <> "I" and isvowel <> "O" and isvowel <> "U") Then
			cons = isvowel
			Exit For
		End If
	Next
End Function

Private Function getPath()
	With CreateObject("WScript.Shell")
		getPath = .CurrentDirectory
	End With 
End Function
