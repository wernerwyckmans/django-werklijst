
import openpyxl as xl
from openpyxl import Workbook
import re
import itertools
import datetime
# TODO: alles omwerken naar dictionary !

MAX_LIST = 0
MIN_STAFF = 0
MAX_STAFF = 0


def main(wb, weekkeuze):

    control_data = list()
    weken = kiesweken(wb, weekkeuze)
    li = maaklijsten(wb)
    dagschemas = maakalledagschemas(weken, li)

    # voer de verschillende controles uit
    control_data += controleeraflossing(dagschemas, li)
    control_data += controleerchefok(dagschemas)
    control_data += checkdaywithoutfunction(dagschemas)
    control_data += checkwachten(weken, li)
    control_data += checkrecup(dagschemas)
    control_data += check_v3(dagschemas)

    return control_data


def kiesweken(workbook, weekkeuze):
    sheets = [workbook['WEEK1'], workbook['WEEK2'], workbook['WEEK3'], workbook['WEEK4']]

    if weekkeuze == 'all':
        sheets_to_use = sheets
    else:
        sheets_to_use = [workbook[weekkeuze]]
    return sheets_to_use


def maaklijsten(workbook):
    # maakt de lijsten aan om in andere functies te gebruiken, output is een lijst van lijsten
    # TODO: functie herschrijven, alle lijsten zullen beschikbaar zijn op afzonderlijk tabblad
    # TODO: Ook hier te herwerken naar dictionary!
    # lijsten[0] = staf, lijsten[1] = Intensieve, lijsten[2] = posities virga
    lijsten = []
    staff = []
    sheet = workbook['WEEK4']

    posities_virga = []
    posities_salvator = []
    setdimensions(sheet)

    # Maak lijst stafleden
    for i in range(MIN_STAFF, MAX_STAFF + 1):  # 131-157, dit mogen geen literals zijn !
        if sheet.cell(i, 3).value is not None:
            staff.append(sheet.cell(i, 3).value)
    lijsten.append(staff)

    # Maak lijst Intensivisten
    # sheet = workbook['Lijsten']  # Opgepast, in werklijst Jeroen bestaat dit tabblad nog niet !!!
    # for i in range(2, len(sheet['A']) + 1):
    #     if sheet.cell(i, 1).value is not None:
    #         intensieve.append(sheet.cell(i, 1).value)
    intensieve = ["BRN", "BMI", "DAE", "DPI", "DBS", "GRE", "GIN", "HBT", "HRS", "HEJ", "JAH", "JMP", "NLM", "STB", "SWV", "VAA", "VTM", "VDB", "WKW", ]
    lijsten.append(intensieve)

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


def maakalledagschemas(sheets_to_use, lijsten):
    # staff is eerste lijst van lijsten, later kunnen andere lijsten toegevoegd worden
    # return is een lijst van alle dagschema's van alle stafleden over gekozen periode
    staff = lijsten[0]
    cumulschemas = []

    for sheet in sheets_to_use:
        setdimensions(sheet)
        for staflid in staff:
            cumulschemas.append(maakindividueledagschemas(staflid, sheet))
    flattedlijst = list(itertools.chain.from_iterable(cumulschemas))

    return flattedlijst


def maakindividueledagschemas(staflid, sheet):
    # genereer een dagschema per staflid per dag
    # return een lijst met dagschemas van staflid

    individueelschema = []
    # doorloop alle dagen van een week
    for col in range(3, 10):

        # datum in rij 4
        # date = sheet.cell(4, col).value.date()
        temp_date = sheet.cell(4, col).value.date()
        date = temp_date.strftime("%d/%m/%Y")

        # initieer lijst met naam staflid, dag en datum
        schema = [staflid, sheet.cell(3, col).value, date]

        # loop over ale verschillende posities en diensten tot aan MAX_LIST
        for row in range(3, MAX_LIST + 1):  # 127

            # als staflid gevonden wordt dan staat zijn dienst in kolom 2
            if staflid in str(sheet.cell(row, col).value):  # conversie naar string -> iterable
                schema.append(sheet.cell(row, 2).value)

        # voeg dagschema toe aan individueel schema staflid
        individueelschema.append(schema)

    return individueelschema


def setdimensions(sheet):
    # controleer sheet, zoek tot waar werklijst loopt en van waar tot waar de stafleden staan(BRN-WKW)
    # MAX_LIST, MIN_STAFF en MAX_STAFF worden hier bepaald
    global MAX_LIST
    global MIN_STAFF
    global MAX_STAFF
    for row in range(1, sheet.max_row):
        if sheet.cell(row, 2).value == "RECUP":
            MAX_LIST = row
        if sheet.cell(row, 3).value == "BRN" and row > MAX_LIST + 1:
            MIN_STAFF = row
        if sheet.cell(row, 3).value == "WKW" and row > MAX_LIST:
            MAX_STAFF = row
    print(MAX_LIST, MIN_STAFF, MAX_STAFF)


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


def lookformatches(schema, matchcount, lijst, functie):
    result = ""
    matches = [x for x in schema[3:] if str(x) in lijst]
    if len(matches) > matchcount:
        zalen = " ,".join(matches)
        result += f'{functie} niet compatibel met zaal: {schema[0]} op {schema[1]} {schema[2]} ' + zalen

    return result


def controleeraflossing(schemas, lijsten):
    # komt staflid meerdere keren op zelfde afloslijst voor, of op twee afloslijsten
    result = list()
    for schema in schemas:

        # regex: start met V of S gevolgd door één of meer getallen of een woord (bv V1 inslapend)
        # meerdere keren op afloslijst
        pattern = re.compile(r'^[VS]\d+\w*')
        matches = [x for x in schema[3:] if pattern.match(str(x))or x == 'Eerst Vertrekkende']
        if len(matches) > 1:
            posities = " ,".join(matches)
            result.append(f'Meer dan één keer op afloslijsten: {schema[0]} op {schema[2]}: ' + posities)

        # staflid op afloslijst Salvator maar ingepland Virga Jesse of allen op afloslijst Salvator
        pattern = re.compile(r'^[S]\d\w*')
        matches = [x for x in schema[3:] if pattern.match(str(x)) and (schema[3] in lijsten[2] or len(schema) == 4)]
        if len(matches) >= 1:
            posities = " ,".join(matches)
            result.append(f'Controleer aflossing/zaaltoewijzing voor: {schema[0]} op {schema[2]}: ' + posities)

        # staflid op afloslijst Virga Jesse maar ingepland op Salvator of alleen op afloslijst Virga Jesse
        pattern = re.compile(r'^[V]\d\w*')
        matches = [x for x in schema[3:] if (pattern.match(str(x)) or x == 'Eerst Vertrekkende') and (schema[3] in lijsten[3] or len(schema) == 4)]
        if len(matches) >= 1:
            posities = " ,".join(matches)
            result.append(f'Controleer aflossing/zaaltoewijzing voor: {schema[0]} op {schema[2]}: ' + posities)
        if len(result) == 0:
            result.append('Geen afwijking gevonden in combinatie zaal/aflos')
    return result


def controleerchefok(schemas):
    # chef ok kan niet in bepaalde zalen staan(notcompatible)
    result = list()

    for schema in schemas:
        if "Corona-Coördinator OK VJ" in schema:
            notcompatible = ["ZAAL 11", "ZAAL 12", "ITE 1", "ITE 2", "ITE 3", "ZAAL 7 CARDIO", "HYBRIDE", "EFO 2",
                             "Radiologie 1"]
            match = lookformatches(schema, 0, notcompatible, "Corona-Coördinator OK VJ")
            if match != "":
                result.append(match)

        if "CHEF OK" in schema:
            notcompatible = ["OFTALMO ZAAL 11", "OFTALMO ZAAL 12", "S-ZAAL 7", "S-ZAAL 8", "S-ZAAL 9"]
            match = lookformatches(schema, 0, notcompatible, "CHEF OK")
            if match != "":
                result.append(match)
        if len(result) == 0:
            result.append("Corona-Coördinator OK VJ en Chef OK Salvator in correcte zalen")
    return result


def checkwachten(weken, lijsten):
    result = list()
    intensieve = lijsten[1]
    for sheet in weken:
        for col in range(3, 10):
            date = sheet.cell(4, col).value.date()
            wachten = {}
            for row in range(40, 50):  # Virga Jesse

                # maak dict wachten met key(functie) uit col 2 en value (staflid) uit col en row
                wachten[sheet.cell(row, 2).value] = sheet.cell(row, col).value

            for row in range(120, 130):  # Salvator

                # maak dict wachten met key(functie) uit col 2 en value (staflid) uit col en row
                wachten[sheet.cell(row, 2).value] = sheet.cell(row, col).value

            if not (wachten['V1 (Coronapermanentie)'] in intensieve or wachten['V2 (INSLAPEND)'] in intensieve or wachten['S1 (Coronapermanentie)'] in intensieve):

                result.append(f'Check Wachten op {date}')
    if len(result) == 0:
        result.append('Geen problemen gevonden in de wachtsamenstelling')
    return result


def checkrecup(schemas):
    result = list()
    for schema in schemas:
        if ("recup" in schema or "RECUP" in schema) and not checkweekend(schema[1]):
            if len(schema) > 4:
                result.append(f'Controleer recuperatie voor: {schema[0]} op {schema[2]}')
    if len(result) == 0:
        result.append('Recuperaties in orde')
    return result


def check_v3(schemas):
    # heeft een staflid dag na V3 een V16, snipperdag of onverwachte snipperdag
    result = list()
    for i in range(0, len(schemas)):
        if "V3 (THUISWACHT)" in schemas[i]:

            # op zaterdag en zondag zijn V16, snipperdag of onverwachte snipperdag niet mogelijk
            if not (schemas[i][1] == 'Vrijdag' or schemas[i][1] == 'Zaterdag'):
                if not ("V16" in schemas[i + 1] or "SNIPPERDAG" in schemas[i + 1] or "Onverwachte Snipperdag" in schemas[i + 1]):
                    result.append(f'Controleer positie na V3 van {schemas[i][0]} op {schemas[i][2]}')
    if len(result) == 0:
        result.append('Posities na V3 in orde.')
    return result









