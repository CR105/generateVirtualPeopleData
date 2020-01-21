Option Explicit
Dim objexcelW, objWorkbookW, objDriverSheetW
Dim getData
Dim Iterator

set objexcelW = Createobject("Excel.Application")
Set objWorkbookW = objExcelW.WorkBooks.Add(getPath & "\nombresApellidos_GEN.xlsx")
Set objDriverSheetW = objWorkbookW.Worksheets("Names")
objexcelW.Visible = False


For Iterator = 1 To 11 Step 1
	Set getData = getVirtualPersonData()
	objDriverSheetW.cells(Iterator+1, 1).Value = getData("Nombre")
	objDriverSheetW.cells(Iterator+1, 2).Value = getData("ApellidoPaterno")
	objDriverSheetW.cells(Iterator+1, 3).Value = getData("ApellidoMaterno")
	objDriverSheetW.cells(Iterator+1, 4).Value = getData("FechaNacimiento")
	objDriverSheetW.cells(Iterator+1, 5).Value = getData("CiudadNacimiento")
	objDriverSheetW.cells(Iterator+1, 6).Value = getData("Genero")
	objDriverSheetW.cells(Iterator+1, 7).Value = getData("CURP")
	objDriverSheetW.cells(Iterator+1, 8).Value = getData("RFC")
Next
objexcelW.ActiveWorkbook.SaveAs getPath & "\nombresApellidos_GEN1.xlsx"
objexcelW.ActiveWorkbook.close

'**********************************************************
'@author Carlos Ríos
'@description Generate virtual persona data, include: Firts name, last name, second name, birthdate, place of birth, gender, CURP and RFC
'**********************************************************
Public Function getVirtualPersonData()
	Dim intGenero, intCiudad
	Dim objData, objDicGen, objDicCity
	Dim objexcelR, objWorkbookR, objDriverSheetR
	Dim numNames(2)
	Dim numApell
	Set objData = createobject("Scripting.Dictionary")
	Set objDicGen = createobject("Scripting.Dictionary")
	Set objDicCity = createobject("Scripting.Dictionary")
	set objexcelR = Createobject("Excel.Application")
	set objWorkbookR = objExcelR.WorkBooks.Open(getPath & "\nombresApellidos.xlsx")
	Set objDriverSheetR = objWorkbookR.Worksheets("Names")
	objexcelR.Visible = False

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

	numNames(1) = (objDriverSheetR.Cells(objDriverSheetR.Rows.Count, 2).End(-4162).Row)-1
	numNames(2) = (objDriverSheetR.Cells(objDriverSheetR.Rows.Count, 1).End(-4162).Row)-1
	numApell = (objDriverSheetR.Cells(objDriverSheetR.Rows.Count, 3).End(-4162).Row)-1

	intGenero = RandomNumber(1,2)
	intCiudad = RandomNumber(1,32)
	objData.add "ApellidoPaterno", trim(objDriverSheetR.cells(RandomNumber(2,numApell), 3))
	objData.add "ApellidoMaterno", trim(objDriverSheetR.cells(RandomNumber(2,numApell), 3))
	objData.add "Nombre", trim(objDriverSheetR.cells(RandomNumber(2,numNames(intGenero)) , intGenero))
	objData.add "FechaNacimiento", Cstr(RandomNumber(1959,1999)) & "/" & Left("0" + Cstr(RandomNumber(1,12)), 2) & "/" & Left("0" + Cstr(RandomNumber(1,28)), 2)
	objData.add "Genero", objDicGen.Item(Cstr(intGenero))
	objData.add "CiudadNacimiento", Cstr(objDriverSheetR.cells(intCiudad+1, 4))
	objData.add "CURP", getCURP(objData("Nombre"), objData("ApellidoPaterno"), objData("ApellidoMaterno"), objData("FechaNacimiento"), objData("Genero"), objDicCity.Item(Cstr(intCiudad)))
	objData.add "RFC", getRFC(objData("Nombre"), objData("ApellidoPaterno"), objData("ApellidoMaterno"), objData("FechaNacimiento"))
	objexcelR.ActiveWorkbook.close
	Set getVirtualPersonData = objData
	Set objData = Nothing
	Set objDicGen = Nothing
	Set objDicCity = Nothing
	set objexcelR = Nothing
	set objWorkbookR = Nothing
	Set objDriverSheetR = Nothing
End Function

'**********************************************************
'@author Carlos Ríos
'@description Generate a random number between of two value, min and max
'**********************************************************
Public Function RandomNumber(nLow, nHigh)
	' Dim objRandom
	' WScript.Sleep 002
	' Set objRandom = CreateObject("System.Random")
	' RandomNumber = objRandom.Next_2(nLow, nHigh)
	RandomNumber = Int((nHigh-nLow+1)*Rnd+nLow)
End Function

'**********************************************************
'@author Carlos Ríos
'@description Get the CURP of the virtual person
'**********************************************************
Public Function getCURP(firtsName, lastName, OtherName, birthDate, gender, city)
	Dim preCURP
	preCURP = Mid(lastName, 1, 2) & Mid(OtherName, 1, 1) & Mid(firtsName, 1, 1) & Mid(birthDate, 3, 2) & Mid(birthDate, 6, 2) & Mid(birthDate, 9, 2) & gender & city & cons(lastName) & cons(OtherName) & cons(firtsName) & "0" & RandomNumber(1, 9)
	getCURP = Ucase(preCURP)
End Function

'**********************************************************
'@author Carlos Ríos
'@description Get the RFC of the virtual person
'**********************************************************
Public Function getRFC(firtsName, lastName, OtherName, birthDate)
	Dim preRFC
	Dim lstAlfa
	lstAlfa = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
	preRFC = Mid(lastName, 1, 2) & Mid(OtherName, 1, 1) & Mid(firtsName, 1, 1) & Mid(birthDate, 3, 2) & Mid(birthDate, 6, 2) & Mid(birthDate, 9, 2) & Mid(lstAlfa, RandomNumber(1,36), 1) & Mid(lstAlfa, RandomNumber(1,36), 1) & Mid(lstAlfa, RandomNumber(1,36), 1)
	getRFC = Ucase(preRFC)
End Function

'**********************************************************
'@author Carlos Ríos
'@description Search the firts vowel after of the firts character in a string
'**********************************************************
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

'**********************************************************
'@author Carlos Ríos
'@description Get a path where execute script for relative path
'**********************************************************
Private Function getPath()
	With CreateObject("WScript.Shell")
		getPath = .CurrentDirectory
	End With 
End Function
