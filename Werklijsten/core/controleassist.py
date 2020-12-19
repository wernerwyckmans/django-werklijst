import re
import itertools

MAX_LIST = 0
MIN_STAFF = 0
MAX_STAFF = 0
MAX_VIRGA = 0
MIN_SAL = 0
MAX_SAL = 0


def main(wb, weekkeuze):

    control_data = list()
    weken = kiesweken(wb, weekkeuze)
    li = maaklijsten(wb)
    dagschemas = maakalledagschemas(weken, li)

    # voer de verschillende controles uit
    control_data += checkdaywithoutfunction(dagschemas)
    control_data += checkrecup(dagschemas)

    return control_data


def setdimensions(sheet):
    """Bepaal de dimensies van het werkblad met de cruciale posities"""
    # controleer sheet, zoek tot waar werklijst loopt en van waar tot waar de stafleden staan(BRN-WKW)
    # MAX_LIST, MIN_STAFF en MAX_STAFF worden hier bepaald
    # zoek wat de min en max rijnummers voor posities in virga en salvator zijn
    global MAX_LIST
    global MIN_STAFF
    global MAX_STAFF
    global MAX_VIRGA
    global MIN_SAL
    global MAX_SAL
    for row in range(1, sheet.max_row):
        if sheet.cell(row, 2).value == "RECUP":
            MAX_LIST = row
        if sheet.cell(row, 2).value == "Radiologie 2":
            MAX_VIRGA = row
        if sheet.cell(row, 2).value == "S-ITE 1 Corona-Coordinator Sal":
            MIN_SAL = row
        if sheet.cell(row, 2).value == "RAADPLEGING Salvator":
            MAX_SAL = row
        if sheet.cell(row, 3).value == "BRN" and row > MAX_LIST + 1:
            MIN_STAFF = row
        if sheet.cell(row, 3).value == "WKW" and row > MAX_LIST:
            MAX_STAFF = row


def maaklijsten(workbook):
    # bijkomende functie maken !
    # maakt de lijsten aan om in andere functies te gebruiken, output is een lijst dictionaries
    # lijsten[0] = staff, lijsten[1]= intensieve, lijsten[2]= posities virga, lijsten[3] = posities salvator
    lijst = {"staff": [], "intensievist": [], "posities_virga": [], "posities_salvator": [], "cardio_anesthesist": [],
             "assistent_anesthesie": [], "assistent_ite": []}

    # Maak lijst Intensivisten
    sheet = workbook['ARTSEN']
    for i in range(2, len(sheet['B']) + 1):
        if sheet.cell(i, 2).value is not None:
            lijst["intensievist"].append(sheet.cell(i, 2).value)

    # Maak lijst stafleden
    for i in range(2, len(sheet['A']) + 1):
        if sheet.cell(i, 1).value is not None:
            lijst["staff"].append(sheet.cell(i, 1).value)

    # Maak lijst cardio_anesthesist
    for i in range(2, len(sheet['C']) + 1):
        if sheet.cell(i, 3).value is not None:
            lijst["cardio_anesthesist"].append(sheet.cell(i, 3).value)

    # Maak lijst assistenten_anesthesie
    for i in range(2, len(sheet['G']) + 1):
        if sheet.cell(i, 7).value is not None:
            lijst["assistent_anesthesie"].append(sheet.cell(i, 7).value)

    # Maak lijst assistent_ite
    for i in range(2, len(sheet['H']) + 1):
        if sheet.cell(i, 8).value is not None:
            lijst["assistent_ite"].append(sheet.cell(i, 8).value)

    # Maak een lijst van Virga Jesse posities
    sheet = workbook['WEEK4']
    setdimensions(sheet)
    for i in range(5, MAX_VIRGA + 1):
        if sheet.cell(i, 2).value is not None:
            lijst["posities_virga"].append(sheet.cell(i, 2).value)

    # Maak lijst van Salvator posities
    for i in range(MIN_SAL, MAX_SAL + 1):
        if sheet.cell(i, 2).value is not None:
            lijst["posities_salvator"].append(sheet.cell(i, 2).value)
    # print(lijst["posities_virga"])

    return lijst


def kiesweken(workbook, weekkeuze):
    sheets = [workbook['WEEK1'], workbook['WEEK2'], workbook['WEEK3'], workbook['WEEK4']]

    if weekkeuze == 'all':
        sheets_to_use = sheets
    else:
        sheets_to_use = [workbook[weekkeuze]]
    return sheets_to_use


def maakindividueledagschemas(assist, sheet):
    # genereer een dagschema per assistent per dag
    # return een lijst met dagschemas van assistent

    individueelschema = []
    # doorloop alle dagen van een week
    for col in range(3, 10):

        # datum in rij 4
        # date = sheet.cell(4, col).value.date()
        temp_date = sheet.cell(4, col).value.date()
        date = temp_date.strftime("%d/%m/%Y")

        # initieer lijst met naam assist, dag en datum
        schema = [assist, sheet.cell(3, col).value, date]

        # loop over ale verschillende posities en diensten tot aan MAX_LIST
        for row in range(3, MAX_LIST + 1):  # 127

            # als staflid gevonden wordt dan staat zijn dienst in kolom 2
            # if assist in str(sheet.cell(row, col).value):  # conversie naar string -> iterable
            if re.search(r'\b' + assist + r'\b', str(sheet.cell(row, col).value)):
                schema.append(sheet.cell(row, 2).value)

        # voeg dagschema toe aan individueel schema staflid
        individueelschema.append(schema)

    return individueelschema


def maakalledagschemas(sheets_to_use, lijsten):
    # assist is eerste lijst van lijsten, later kunnen andere lijsten toegevoegd worden
    # return is een lijst van alle dagschema's van alle assistenten over gekozen periode
    assistenten = lijsten["assistent_anesthesie"] + lijsten["assistent_ite"]
    cumulschemas = []

    for sheet in sheets_to_use:
        setdimensions(sheet)
        for assist in assistenten:
            cumulschemas.append(maakindividueledagschemas(assist, sheet))
    flattedlijst = list(itertools.chain.from_iterable(cumulschemas))

    return flattedlijst


def checkweekend(day):
    weekend = False
    if day == "Zaterdag" or day == "Zondag":
        weekend = True
    return weekend


def checkdaywithoutfunction(schemas):
    # weekdag zonder toewijzingen
    # TODO: lijst met feestdagen importeren -> moet bestaan in python!
    result = list()
    # schema bestaat altijd uit minstens [staflid, dag, datum ], de volgende elementen zijn toewijzigingen
    for schema in schemas:
        if len(schema) == 3:
            if not checkweekend(schema[1]):
                result.append(f'Geen toewijzingen voor  {schema[0]} op {schema[1]} {schema[2]}')
    return result


def checkrecup(schemas):
    result = list()
    for schema in schemas:
        if ("Recup" in schema or "RECUP ITE ASSISTENT" in schema) and not checkweekend(schema[1]):
            if len(schema) > 4:
                result.append(f'Controleer recuperatie voor: {schema[0]} op {schema[2]}')
    if len(result) == 0:
        result.append('Recuperaties in orde')
    return result
