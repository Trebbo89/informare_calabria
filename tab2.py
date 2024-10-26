import customtkinter as ctk
from functions import *
from constants import *

# metodo per la generazione del tab 1
def tab_2_generation(app, ctk = ctk):
    try:
        
        all_students_course = list()
        
        configuration_row_or_column_for_grid_template(app, 13, True)
        configuration_row_or_column_for_grid_template(app, 11, False)
        
        label_students_file = ctk.CTkLabel(app, text='File studenti:')
        label_students_file.grid(padx=PADDING_X, pady=PADDING_Y, row=4, column=3, sticky='we', columnspan=2)
        
        input_text_filaname = ctk.CTkEntry(app, state='disabled', placeholder_text='Nome file selzionato')
        input_text_filaname.grid(padx=PADDING_X, pady=PADDING_Y, row=4, column=6, sticky='we', columnspan=2)
        
        input_file_get_students = ctk.CTkButton(app, text='Seleziona file studenti corso', command=lambda: operations_input_file_get_students(app, input_text_filaname, input_import_mynaparrot_report, all_students_course))
        input_file_get_students.grid(padx=PADDING_X, pady=PADDING_Y, row=5, column=3, sticky='we', columnspan=5)
        
        input_import_mynaparrot_report = ctk.CTkButton(app, text='Seleziona file report mynaparrot', state='disabled', command=lambda: operations_import_report_mynaparrot(app, all_students_course))
        input_import_mynaparrot_report.grid(padx=PADDING_X, pady=PADDING_Y, row=6, column=3, sticky='we', columnspan=5)
    except Exception as e:
        generate_modal(app, TEXT_MESSAGGIO_ERRORE)
        generate_log_errore(app, e)