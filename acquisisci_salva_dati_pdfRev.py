import os
import PyPDF2
import csv
import xml.etree.ElementTree as ET

# Funzione per acquisire i valori dei campi da un singolo PDF
def acquisisci_valori_dei_campi(pdf_file, xml_file, campi_interessati):
    pdf = PyPDF2.PdfFileReader(open(pdf_file, 'rb'))
    root = ET.Element('pdf_structure')
    for page_num in range(pdf.getNumPages()):
        page = pdf.getPage(page_num)
        try:
            annotations = page['/Annots']
        except KeyError:
            continue
        for annotation in annotations:
            field_info = ET.SubElement(root, 'field')
            field_info.set('page', str(page_num + 1))
            for key, value in annotation.getObject().items():
                field_info.set(key[1:], str(value))
    tree = ET.ElementTree(root)
    tree.write(xml_file, encoding='utf-8', xml_declaration=True)
    tree = ET.parse(xml_file)
    root = tree.getroot()
    valori_dei_campi = {}
    for campo in campi_interessati:
        elemento = root.find(".//field[@T='%s']" % campo)
        if elemento is not None:
            valore = elemento.get('V')
            valori_dei_campi[campo] = valore
    return valori_dei_campi

# Cartella con i file PDF da elaborare
cartella_pdf = 'pdf_da_acquisire'
# Cartella di output per i file CSV
cartella_csv = 'csv_da_salvare'
# Campi di interesse
campi_interessati = [
    'cognome_nome',
    'luogo_nascita',
    'data_nascita',
    'plesso',
    'qualifica',
    'richiesta_giorni',
    'giorno_a',
    'giorno_b',
    'giorno_c',
    'giorno_d',
    'giorno_e',
    'giorno_f',
    'Tot_ore_fatte',
    'tot_ore_da_fare',
    'protocollo',
    'data',
]

# Assicurati che la cartella di output per i file CSV esista
if not os.path.exists(cartella_csv):
    os.makedirs(cartella_csv)

# Inizializza il file CSV per l'output
csv_file = os.path.join(cartella_csv, 'output.csv')
with open(csv_file, 'w', newline='') as csvfile:
    csv_writer = csv.DictWriter(csvfile, fieldnames=campi_interessati)
    csv_writer.writeheader()

# Ciclo su tutti i file PDF nella cartella
for filename in os.listdir(cartella_pdf):
    if filename.endswith('.pdf'):
        pdf_file = os.path.join(cartella_pdf, filename)
        xml_file = os.path.join(cartella_csv, f'{os.path.splitext(filename)[0]}.xml')
        # Acquisisci i valori dei campi
        valori_dei_campi = acquisisci_valori_dei_campi(pdf_file, xml_file, campi_interessati)
        # Scrivi i valori nel file CSV
        with open(csv_file, 'a', newline='') as csvfile:
            csv_writer = csv.DictWriter(csvfile, fieldnames=campi_interessati)
            csv_writer.writerow(valori_dei_campi)

print(f"Dati estratti da file PDF e salvati in {csv_file}")
