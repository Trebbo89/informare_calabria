import customtkinter as ctk
from functions import *
from constants import *

# metodo per la generazione del tab 1
def tab_1_generation(app, ctk = ctk):
    try:
        course_dict_with_students = dict()
        
        configuration_row_or_column_for_grid_template(app, 6, True)
        configuration_row_or_column_for_grid_template(app, 3, False)
        
        textarea_corso = ctk.CTkTextbox(app);
        textarea_corso.grid(padx=PADDING_X, pady=PADDING_Y, row=0, column=2, sticky='wens', rowspan=6)
        textarea_corso.configure(state='disabled')
        
        course_input = ctk.CTkEntry(app, placeholder_text='Nome corso')
        course_input.grid(padx=PADDING_X, pady=PADDING_Y, row=0, column=0, sticky='we', columnspan=2)
        
        student_input_name_lastname = ctk.CTkEntry(app, placeholder_text='Nome Cognome studente')
        student_input_name_lastname.grid(padx=PADDING_X, pady=PADDING_Y, row=1, column=0, sticky='we', columnspan=2)
        
        checkbox_var = ctk.StringVar(value = 'off')
        checkbox_update_student = ctk.CTkCheckBox(app, text='Modifica Studente', variable=checkbox_var, onvalue='on', offvalue='off',command=lambda : operations_flag_checkbox(app, option_id_student, button_add_student, checkbox_var.get(), course_dict_with_students, button_delete_student) )
        checkbox_update_student.grid(padx=PADDING_X, pady=PADDING_Y, row=2, column=0, sticky='we')
        
        option_id_student = ctk.CTkOptionMenu(app, values=[OPTION_ID_STUDENTE_DEFAULT], state='disabled', command=lambda selected_value: enable_or_disable_delete_button(app, button_delete_student, checkbox_var.get(), selected_value))
        option_id_student.grid(padx=PADDING_X, pady=PADDING_Y, row=2, column=1, sticky='we')
        
        button_add_student = ctk.CTkButton(app, text=TEXT_INSERT_STUDENT, command=lambda : operations_insert_update_student(app, ctk, course_dict_with_students, course_input.get(), student_input_name_lastname.get(), textarea_corso, checkbox_var.get(), option_id_student) )
        button_add_student.grid(padx=PADDING_X, pady=PADDING_Y, row=3, column=0, sticky='we')
        
        button_delete_student = ctk.CTkButton(app, text=TEXT_DELETE_STUDENT, state='disabled', command=lambda : delete_student(app, textarea_corso, course_dict_with_students, option_id_student) )
        button_delete_student.grid(padx=PADDING_X, pady=PADDING_Y, row=3, column=1, sticky='we')
        
        button_add_file = ctk.CTkButton(app, text='Importa studenti da file', command=lambda: operations_import_students_from_file(app, textarea_corso, course_dict_with_students, course_input))
        button_add_file.grid(padx=PADDING_X, pady=PADDING_Y, row=4, column=0, sticky='we', columnspan=2)
        
        button_generate_course_file = ctk.CTkButton(app, text='Genera File Corso', command=lambda : generate_course_file(app, course_dict_with_students))
        button_generate_course_file.grid(padx=PADDING_X, pady=PADDING_Y, row=5, column=0, sticky='we', columnspan=2)
    except Exception as e:
        generate_modal(app, TEXT_MESSAGGIO_ERRORE)
        generate_log_errore(app, e)
    
    
    