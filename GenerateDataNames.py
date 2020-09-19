import openpyxl
import random
from openpyxl import Workbook
from datetime import datetime

objDicGen = {"1":["H", "nameMan"], "2":["M","nameWomen"]}
objDicCity = {"1":"AS","2":"BC","3":"BS","4":"CC","5":"CS","6":"CH","7":"CL","8":"CM","9":"DF","10": "DG","11": "GT","12": "GR","13": "HG","14": "JC","15": "MC","16": "MN","17": "MS","18": "NT","19": "NL","20": "OC","21": "PL","22": "QO","23": "QR","24": "SP","25": "SL","26": "SR","27": "TC","28": "TS","29": "TL","30": "VZ","31": "YN","32": "ZS"}

# Read Excel
def getDataXLSFile():
    
    wb = openpyxl.load_workbook('nombresApellidos.xlsx') 
    sheet = wb["Names"]
    
    listLasNam = []
    listNamMan = []
    listNamWom = []
    listEF = []

    colA = sheet['A']
    colB = sheet['B']
    colC = sheet['C']
    colD = sheet['D']

    for val in colA:
        if (val.value != None):
            listNamMan.append(val.value)

    for val in colB:
        if (val.value != None):
            listNamWom.append(val.value)
    
    for val in colC:
        if (val.value != None):
            listLasNam.append(val.value)

    for val in colD:
        if (val.value != None):
            listEF.append(val.value)

    data = {"nameMan":listNamMan,"nameWomen":listNamWom,"lastName":listLasNam,"eF":listEF}
    return data

# Write Excel
def saveXLS(listData):
    now = datetime.now()
    wb = Workbook()
    sheetMain = wb.active
    
    # Headers
    sheetMain.append(["Nombre","Apellido Paterno","Apellido Materno","Fecha de nacimiento","Ciudad","Genero","CURP","RFC"])

    for lList in listData:
        sheetMain.append(lList)
    
    wb.save('results/test' + now.strftime("%Y%m%d-%H%M%S") + '.xlsx') 
    pass

# Get consonant of string
def cons(someString):
    consLet = ''

    for indx in someString[1:]:
        if indx in 'AEIOU':
            pass
        else:
            consLet = indx
            break

    if consLet == '':
        consLet = someString[1] 

    return consLet

# Create CURP
def getCURP(firtsName, fatherlastName, motherLastName, birthDate, gender, city):
    return fatherlastName[0:2] + motherLastName[0] + firtsName[0] + birthDate[2:4] + birthDate[5:7] + birthDate[8:10] + gender + city + cons(fatherlastName) + cons(motherLastName) + cons(firtsName) + "0" + str(random.randint(1, 9))

# Create RFC
def getRFC(firtsName, fatherlastName, motherLastName, birthDate):
    lstAlfa = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    return fatherlastName[0:2] + motherLastName[0] + firtsName[0] + birthDate[2:4] + birthDate[5:7] + birthDate[8:10] + lstAlfa[random.randint(1, len(lstAlfa)-1)] + lstAlfa[random.randint(1, len(lstAlfa)-1)] + lstAlfa[random.randint(1, len(lstAlfa)-1)]

numVirtualUsers = 50
vUser = []
print("Loading...")
virtualData = getDataXLSFile()

for iter in range(0, numVirtualUsers):
    intGenero = random.randint(1, 2)
    intCiudad = random.randint(1, 32)
    genero = objDicGen[str(intGenero)][0]
    cdgCiudadNacimiento = objDicCity[str(intCiudad)]
    fechaNacimiento = str(random.randint(1950, 2001)) + "/" + str(random.randint(1, 12)).zfill(2) + "/" + str(random.randint(1, 28)).zfill(2)
    name = virtualData[objDicGen[str(intGenero)][1]][random.randint(1, len(virtualData[objDicGen[str(intGenero)][1]])-1)]
    fatherLastName = virtualData["lastName"][random.randint(1, len(virtualData["lastName"])-1)]
    motherLastName = virtualData["lastName"][random.randint(1, len(virtualData["lastName"])-1)]
    ciudadNacimiento = virtualData["eF"][intCiudad]
    RFC = getRFC(name, fatherLastName, motherLastName, fechaNacimiento)
    CURP = getCURP(name, fatherLastName, motherLastName, fechaNacimiento, genero, cdgCiudadNacimiento)
    vUser.append([name, fatherLastName, motherLastName, fechaNacimiento, ciudadNacimiento, genero, CURP, RFC])
    pass

saveXLS(vUser)
print("...Finish.")
