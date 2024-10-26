import customtkinter as ctk
from functions import *
from constants import *
from tab1 import tab_1_generation
from tab2 import tab_2_generation

app = ctk.CTk()
app.geometry('700x500')
app.title('Gestione Report Excel')

configuration_row_or_column_for_grid_template(app, 1, False)

# Generazione TAB
tabview = ctk.CTkTabview(app)
tabview.pack(fill="both", expand=True)
    
for tab_name in TAB_NAME_LIST:
    tabview.add(tab_name)
    
tabview.set(TAB_NAME_LIST[0])

tab_1_generation(tabview.tab(TAB_NAME_LIST[0]), ctk)
tab_2_generation(tabview.tab(TAB_NAME_LIST[1]), ctk)

app.mainloop()