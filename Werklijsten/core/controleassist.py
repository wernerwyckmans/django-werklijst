import re
import itertools

MAX_LIST = 0
MIN_STAFF = 0
MAX_STAFF = 0


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
    # controleer sheet, zoek tot waar werklijst loopt en van waar tot waar de stafleden staan(BRN-WKW)
    # MAX_LIST, MIN_STAFF en MAX_STAFF worden hier bepaald
    global MAX_LIST
    global MIN_STAFF
    global MAX_STAFF
    for row in range(1, sheet.max_row):
        if sheet.cell(row, 2).value == "RECUP":
            MAX_LIST = row + 1
        if sheet.cell(row, 3).value == "BRN" and row > MAX_LIST + 1:
            MIN_STAFF = row
        if sheet.cell(row, 3).value == "WKW" and row > MAX_LIST:
            MAX_STAFF = row
    print(MAX_LIST, MIN_STAFF, MAX_STAFF)


def maaklijsten(workbook):
    # maakt de lijsten aan om in andere functies te gebruiken, output is een lijst van lijsten
    # TODO: functie herschrijven, alle lijsten zullen beschikbaar zijn op afzonderlijk tabblad
    # Ook hier te herwerken naar dictionary!
    # lijsten[0] = staf, lijsten[1] = Intensieve, lijsten[2] = posities virga
    lijsten = []
    assist = []
    sheet = workbook['WEEK4']
    posities_virga = []
    posities_salvator = []
    setdimensions(sheet)

    # Maak lijst assistenten
    for i in range(MAX_STAFF + 2, MAX_STAFF + 28):  # 131-157, dit mogen geen literals zijn !
        if sheet.cell(i, 3).value is not None:
            assist.append(sheet.cell(i, 3).value)
    lijsten.append(assist)

    # Maak een lijst van Virga Jesse posities
    sheet = workbook['WEEK4']
    for i in range(5, 35):  # TODO Nog aanpassen via setdimensions
        if sheet.cell(i, 2).value is not None:
            posities_virga.append(sheet.cell(i, 2).value)
    lijsten.append(posities_virga)

    # Maak lijst van Salvator posities
    for i in range(101, 122):  # TODO: Nog aanpassen via setdimensions
        if sheet.cell(i, 2).value is not None:
            posities_salvator.append(sheet.cell(i, 2).value)
    lijsten.append(posities_salvator)

    return lijsten


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
    assistenten = lijsten[0]
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
