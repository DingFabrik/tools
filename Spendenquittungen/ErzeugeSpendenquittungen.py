import pandas as pd # pandas selbst
import numpy as np # scientific computing

import jinja2
import os
import codecs

# import matplotlib.pyplot as plt # Grafiken

def cleanPLZ(stringToClean):
    return stringToClean.replace('.0','')

# http://chrisalbon.com/python/pandas_list_unique_values_in_column.html
# Set ipython's max row display
pd.set_option('display.max_row', 10000)
# Set iPython's max column width to 50
pd.set_option('display.max_columns', 50)
# A set number format to 2 digits
pd.set_option('display.float_format', lambda x: '%.2f' % x)
# http://stackoverflow.com/questions/20625582/how-to-deal-with-this-pandas-warning
# komische Fehlermeldung beim Drop von Spalten loswerden
pd.options.mode.chained_assignment = None  # default='warn'

##################### Prepare Stammdaten
# lade Daten
stammdaten = pd.read_excel('Stammdaten.xlsx')
# Remove NaN values by " foc/or strings
stammdaten['Vorname'].fillna(value='',inplace=True)
stammdaten['Nachname'].fillna(value='',inplace=True)
#stammdaten['c/o'].fillna(value='',inplace=True)
stammdaten['Straße'].fillna(value='',inplace=True)
stammdaten['PLZ'].fillna(value='',inplace=True)
stammdaten['Ort'].fillna(value='',inplace=True)
#stammdaten['E-Mail'].fillna(value='',inplace=True)
stammdaten['Mitgliedsart'].fillna(value='',inplace=True)

# convert PLZ to string
# apply str function first, then run cleanPLZ on the string
stammdaten['PLZ']= stammdaten.PLZ.apply(str)
stammdaten['PLZ']= stammdaten.PLZ.apply(cleanPLZ)

stammdaten['Suppress'] = ''

# use ID as index
#stammdaten = stammdaten.set_index("ID")

# entferne ehemalige Mitglieder
stammdaten  = stammdaten[stammdaten.Mitgliedsart != 'E']


################################ Prepare Buchungen
# lade die Buchungen
buchungen = pd.read_excel('Buchungen.xlsx', converters={'Klasse' : str, 'Vorgang' : str, })
# Explizite Konvertierung in Strings und Float für zwei Spalten, sonst Probleme beim Matching und Zerlegen
buchungen[['Klasse']] = buchungen[['Klasse']].astype(str)
buchungen[['Betrag']] = buchungen[['Betrag']].astype(float)

# Change format of 'Buchungstag' to datetime
buchungen['Buchungstag'] = pd.to_datetime(buchungen['Buchungstag'],dayfirst=True)

# Remove NaN values by " for strings or 0 for numbers 
buchungen['Vorgang'].fillna(value=0,inplace=True)
buchungen['Empfänger'].fillna(value='',inplace=True)
buchungen['Verwendungszweck'].fillna(value='',inplace=True)
buchungen['Kategorie'].fillna(value='',inplace=True)
buchungen['Klasse'].fillna(value='',inplace=True)
buchungen['Relevant']=buchungen.Kategorie.str.match('^Aufnahmegebühr_2150|Zweckspende_3220|Mitgliedsbeitrag_2110|Spende_3220$')


# entferne irrelevante Buchungen
buchungen  = buchungen[buchungen.Relevant != False]

# Eine Funktion, die die Adresse vorbereitet
# keine überflüssigen Leerzeichen, wenn Feld nicht gefüllt ist
def prepareAddress(id, vorname, name, strasse, plz, ort):
    address = '' # + str(id) + ': '
    if len(vorname)==0:
        address = address + name
    else:
        address = address + vorname + ' ' + name
    if len(strasse)>0:
        address = address + ", " + strasse    
    if len(plz)>0:
        address = address + ", " + plz + ' ' + ort
    return address

# Zerlege die Gesamtsumme in einzelne Bestandteile, um Zahlwort auszugeben
# Siehe http://www.steuer-schutzbrief.de/fileadmin/downloads/BMF-Schreiben/BMF-Schreiben-Zuwendungsbestaetigung-2012-08-30.pdf
def kardinal(summenstring,separator,indicator):
    #print('summenstring',summenstring)
    zahlen = {"1" : "Eins", "2":"Zwei", "3":"Drei", "4":"Vier","5":"Fünf","6":"Sechs","7":"Sieben","8":"Acht","9":"Neun","0":"Null"}
    zahlwort = ''
    zahl = summenstring.split(',')[0]
    for i in zahl:
        zahlwort = zahlwort + zahlen[i]+ separator
    return indicator + separator + zahlwort + indicator

# http://stackoverflow.com/questions/20937538/how-to-display-pandas-dataframe-using-a-format-string-for-columns
#pd.options.display.float_format = '{:,.2f} EUR'.format

class CommaFloatFormatter:
    def __mod__(self, x):
        return str(x).replace('.',',')

latex_jinja_env = jinja2.Environment(
    block_start_string = '\BLOCK{',
    block_end_string = '}',
    variable_start_string = '\VAR{',
    variable_end_string = '}',
    comment_start_string = '\#{',
    comment_end_string = '}',
    line_statement_prefix = '%-',
    line_comment_prefix = '%#',
    trim_blocks = True,
    autoescape = False,
    loader = jinja2.FileSystemLoader(os.path.abspath('.'))
)

# Laden des Templates aus einer Datei
template = latex_jinja_env.get_template('Sammelbestaetigung_Geldzuwendung.tex')

for index, row in stammdaten.iterrows():
    if row["Suppress"] == True:
        pass
    else:
        print(row["ID"])
        address = prepareAddress(row["ID"],row['Vorname'],row['Nachname'],row['Straße'],row['PLZ'],row['Ort'])
    
        beitraege = buchungen[buchungen.Klasse.str.match('^' +  str(row["ID"]) + '$')]
        if len(beitraege) == 0:
            print('Keine Buchungen für', row["ID"])
        else:
            beitraege.drop('Klasse',axis=1,inplace=True)
            beitraege.drop('Verwendungszweck',axis=1,inplace=True)
            beitraege.drop('Relevant',axis=1,inplace=True)
            #beitraege.drop('Jahr',axis=1,inplace=True)
            beitraege.drop('Monat',axis=1,inplace=True)
            beitraege.drop('Empfänger',axis=1,inplace=True)
            beitraege.drop('Konto',axis=1,inplace=True)
            #beitraege.drop('Vorgang',axis=1,inplace=True)
            gesamtsumme = beitraege['Betrag'].sum()
            
            beitraege['Buchungstag'] = beitraege['Buchungstag'].apply(lambda x: x.strftime('%d-%m-%Y'))
            #print(beitraege.to_latex(index=False,float_format=CommaFloatFormatter()))
            texbuchungen = beitraege.applymap(lambda x: str(x).replace('.',',0')).to_latex(index=False)    
            texbuchungen = beitraege.to_latex(index=False)    
            summe = str(gesamtsumme).replace('.',',0') + ' EUR'
            
            dokument = template.render(Spender=address, ID=row['ID'],Summe=summe,kardinal=kardinal(str(summe),'-','xxx'),Buchungen=texbuchungen)
            #print(dokument)
            with codecs.open('./fertig/'+str(row['ID']) + ".tex", "w","utf-8") as letter:
                letter.write(dokument);
                letter.close();
                # Aufruf von pdflatex
                os.system("pdflatex -output-directory=./fertig/ -interaction=batchmode ./fertig/" + str(row['ID']) + ".tex")
        
os.system("del ./fertig/*.log")
os.system("del ./fertig/*.aux")
