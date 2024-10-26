TAB_NAME_LIST = ['Inserimento Corso e Utenti', 'Upload Report Mynaparrot']

PADDING_X = 3
PADDING_Y = 3

# Costanti per file
FILETYPES_TXT = ('text files', '*.txt')
FILETYPES_XLSX = ('xlsx files', '*.xlsx')

IMPORT_STUDENTS_FROM_FILE = 'students_from_file'
IMPORT_REPORT_MYNAPARROT_FILE = 'mynaparrot_file'


# Costanti testi input
TEXT_INSERT_STUDENT = 'Inserimento studente'
TEXT_UPDATE_STUDENT = 'Modifica studente'
TEXT_DELETE_STUDENT = 'Elimina studente'
OPTION_ID_STUDENTE_DEFAULT = 'Seleziona ID studente'

# Costante messaggio di errore
TEXT_MESSAGGIO_ERRORE = 'Si è verificato un errore, consultare i log'

# Costanti per settaggio style excel
BORDER_STYLE_THICK  = 'thick'
BORDER_STYLE_THIN   = 'thin'
BORDER_STYLE_COLOR_BLACK = '000000'

# Costanti generazione excel report
MINUTI_TOLLERANZA_STUDENTE = 15
MINUTI_TOLLERANZA_ACCESSO_AL_CORSO = 30
MINUTI_TOLLERANZA_USCITA_DAL_CORSO = 15

# Costanti testi EXCEL
TEXT_MATERIA = 'Materia'
TEXT_ORE = 'Ore'
TEXT_DOCENTE = 'Docente'

TEXT_UTENTE_ASSENTE = 'ASSENTE'
TEXT_NOTE = 'NOTE'
TEXT_TOTALE_PRESENZE_GIORNALIERE = 'Tot. Presenze giornaliere'
TEXT_TOTALE_ASSENZE_GIORNALIERE  = 'Tot. Assenze giornaliere'
TEXT_TOTALE_ORE_GIORNALIERE = 'Tot. Ore giornaliere'
TEXT_FIRMA_TUTOR = 'Firma tutor'
TEXT_N = 'N. _________________'

NOME_REPORT_GENERATO = 'Report personalizzato.xlsx'

TEXT_REPORT_GENERATO_CON_SUCCESSO = 'Il report è stato generato\ncon successo'

TEXT_NO_BORDER = 'no_border'
TEXT_MERGE = 'merge'

LIST_TEXT_WITHOUT_BORDER = [TEXT_NO_BORDER, TEXT_TOTALE_PRESENZE_GIORNALIERE, TEXT_TOTALE_ASSENZE_GIORNALIERE, TEXT_TOTALE_ORE_GIORNALIERE, TEXT_FIRMA_TUTOR, TEXT_N]

COLUMN_REPORT_EXCEL = ['A', 'B', 'C', 'D']

VALUES_TO_CENTER = [
    TEXT_UTENTE_ASSENTE, 
    TEXT_MATERIA, 
    TEXT_ORE, 
    TEXT_DOCENTE, 
    TEXT_NOTE, 
    TEXT_TOTALE_PRESENZE_GIORNALIERE, 
    TEXT_TOTALE_ASSENZE_GIORNALIERE, 
    TEXT_TOTALE_ORE_GIORNALIERE, 
    TEXT_FIRMA_TUTOR, 
    TEXT_N
]

# Costanti alignment
HORIZONTAL_ALIGNMENT_CENTER     = 'center'
HORIZONTAL_ALIGNMENT_LEFT       = 'left'
HORIZONTAL_ALIGNMENT_RIGHT      = 'right'
HORIZONTAL_ALIGNMENT_JUSTIFY    = 'justify'

VERTICAL_ALIGNMENT_TOP      = 'top'
VERTICAL_ALIGNMENT_CENTER   = 'center'
VERTICAL_ALIGNMENT_BOTTOM   = 'bottom'

# Costanti Date
PATTERN_DD_MM_YY = '%d/%m/%Y'
PATTERN_DD_MM_YY_WITH_TIME = '%d/%m/%Y %H:%M'
PATTERN_DD_MM_YY_WITH_TIME_AND_SECONDS = '%d/%m/%Y %H:%M:%S'