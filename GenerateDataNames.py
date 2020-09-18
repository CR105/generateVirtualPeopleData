import openpyxl
import random
from openpyxl import Workbook


objDicGen = {"1":["H", "nameMan"], "2":["M","nameWomen"]}
objDicCity = {"1":"AS","2":"BC","3":"BS","4":"CC","5":"CS","6":"CH","7":"CL","8":"CM","9":"DF","10": "DG","11": "GT","12": "GR","13": "HG","14": "JC","15": "MC","16": "MN","17": "MS","18": "NT","19": "NL","20": "OC","21": "PL","22": "QO","23": "QR","24": "SP","25": "SL","26": "SR","27": "TC","28": "TS","29": "TL","30": "VZ","31": "YN","32": "ZS"}

#read Excel
def readingXLSFile():
    
    wb = openpyxl.load_workbook('nombresApellidos.xlsx') 
    sheet = wb["Names"]
    
    colA = sheet['A']
    listNamMan = []
    for val in colA:
        if (val.value != None):
            listNamMan.append(val.value)

    colB = sheet['B']
    listNamWom = []
    for val in colB:
        if (val.value != None):
            listNamWom.append(val.value)

    colC = sheet['C']
    listLasNam = []
    for val in colC:
        if (val.value != None):
            listLasNam.append(val.value)

    colD = sheet['D']
    listEF = []
    for val in colD:
        if (val.value != None):
            listEF.append(val.value)

    data = {"nameMan":listNamMan,"nameWomen":listNamWom,"lastName":listLasNam,"eF":listEF}
    return data

#write Excel
def saveXLS():
    
    wb = Workbook()
    # sheetMain = wb.create_sheet("Nombres")
    sheetMain = wb.active
    
    # Headers
    # Nombre | Apellido Paterno | Apellido Materno | Fecha de nacimiento | Ciudad | Genero | CURP | RFC
    
    sheetMain.append(["Nombre","Apellido Paterno","Apellido Materno","Fecha de nacimiento","Ciudad","Genero","CURP","RFC"])

    
    wb.save('results/test.xlsx') 
    pass

def cons(someString):
    return "A"

def getCURP(firtsName, fatherlastName, motherLastName, birthDate, gender, city):
    return fatherlastName[0:2] + motherLastName[0] + firtsName[0] + birthDate[2:4] + birthDate[5:7] + birthDate[8:10] + gender + city + cons(fatherlastName) + cons(motherLastName) + cons(firtsName) + "0" + str(random.randint(1, 9))

def getRFC(firtsName, fatherlastName, motherLastName, birthDate):
    lstAlfa = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    return fatherlastName[0:2] + motherLastName[0] + firtsName[0] + birthDate[2:4] + birthDate[5:7] + birthDate[8:10] + lstAlfa[random.randint(1, len(lstAlfa))] + lstAlfa[random.randint(1, len(lstAlfa))] + lstAlfa[random.randint(1, len(lstAlfa))]


virtualData = readingXLSFile()
intGenero = random.randint(1, 2)
intCiudad = random.randint(1, 32)
genero = objDicGen[str(intGenero)][0]
cdgCiudadNacimiento = objDicCity[str(intCiudad)]
fechaNacimiento = str(random.randint(1950, 2001)) + "/" + str(random.randint(1, 12)).zfill(2) + "/" + str(random.randint(1, 28)).zfill(2)

name = virtualData[objDicGen[str(intGenero)][1]][random.randint(1, len(virtualData[objDicGen[str(intGenero)][1]]))]
fatherLastName = virtualData["lastName"][random.randint(1, len(virtualData["lastName"])-1)]
motherLastName = virtualData["lastName"][random.randint(1, len(virtualData["lastName"])-1)]
ciudadNacimiento = virtualData["eF"][intCiudad]

RFC = getRFC(name, fatherLastName, motherLastName, fechaNacimiento)
CURP = getCURP(name, fatherLastName, motherLastName, fechaNacimiento, genero, cdgCiudadNacimiento)

print("Name:", name)
print("Father Last Name:", fatherLastName)
print("Mother Last Name:", motherLastName)
print("Genero :", genero)
print("Ciudad de Nacimiento:", ciudadNacimiento, " - ",cdgCiudadNacimiento)
print("Fecha de Nacimiento", fechaNacimiento)
print("RFC:", RFC)
print("CURP:", CURP)


# "CURP", getCURP(objData("Nombre"), objData("ApellidoPaterno"), objData("ApellidoMaterno"), objData("FechaNacimiento"), objData("Genero"), objDicCity.Item(Cstr(intCiudad)))

# saveXLS()
