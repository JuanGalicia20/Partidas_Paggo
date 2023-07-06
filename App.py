import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import customtkinter
from PIL import ImageTk, Image
from CTkMessagebox import CTkMessagebox
import os


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light")
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue")




class App(customtkinter.CTk):

    def __init__(self):
        super().__init__()

        customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light")
        customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue")

        self.file1 = ""
        self.file2 = ""
        self.file3 = ""

        self.iconbitmap("imagen_2023-05-12_003813110-removebg-preview.ico")
        self.wm_iconbitmap("imagen_2023-05-12_003813110-removebg-preview.ico")
        self.title("Paggo App Registros Contables")
        self.geometry(f"{1100}x580")

        # configure grid layout (4x4)
        self.grid_columnconfigure(0, weight=0)  # Set weight to 0 for the first column
        self.grid_columnconfigure(1, weight=1)  # Set weight to 1 for the second column

        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure((1, 2, 3), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=100, corner_radius=0)  # Reduce width here
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        # Add image to the left of the label
        self.logo_label = customtkinter.CTkLabel(
            self.sidebar_frame, text="Paggo App", font=customtkinter.CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=5, pady=(20, 10))
        
        # create tabview
        self.tabview = customtkinter.CTkTabview(self.sidebar_frame, width=120, height=500)
        self.tabview.grid(row=2, column=0, padx=5, pady=0)
        self.tabview.add("Cuenta 007")
        self.tabview.add("Otra")
        self.tabview.tab("Cuenta 007").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Otra").grid_columnconfigure(0, weight=1)

        self.file1_button = customtkinter.CTkButton(self.tabview.tab("Cuenta 007"), text="Escoge Archivo\nPortal Paggo",
                                                    command=self.choose_file1)
        self.file1_button.grid(row=2, column=0, padx=5, pady=(50,50))

        self.file2_button = customtkinter.CTkButton(self.tabview.tab("Cuenta 007"), text="Escoge Archivo Bi",
                                                    command=self.choose_file2)
        self.file2_button.grid(row=3, column=0, padx=5, pady=(50,50))
        
        self.file3_button = customtkinter.CTkButton(self.tabview.tab("Cuenta 007"), text="Escoge Archivo\nTransacciones Bi",
                                                    command=self.choose_file3)
        self.file3_button.grid(row=4, column=0, padx=5, pady=(50,50))
        
        
        self.file4_button = customtkinter.CTkButton(self.tabview.tab("Otra"), text="Escoge Archivo\nTransacciones Bi",
                                                    command=self.choose_file3)
        self.file4_button.grid(row=3, column=0, padx=5, pady=(50,50))

        # create Treeview to display DataFrame
        self.dataframe_treeview = ttk.Treeview(self, columns=(), show="headings")
        self.dataframe_treeview.grid(row=0, column=1, padx=10, pady=10, columnspan=3, rowspan=4, sticky="nsew")

        # create horizontal scrollbar for Treeview
        self.dataframe_treeview_scrollbar_x = ttk.Scrollbar(self, orient="horizontal",
                                                            command=self.dataframe_treeview.xview)
        self.dataframe_treeview_scrollbar_x.grid(row=4, column=1, columnspan=3, sticky="ew")
        self.dataframe_treeview.configure(xscrollcommand=self.dataframe_treeview_scrollbar_x.set)

        # create a frame at the bottom
        self.bottom_frame = customtkinter.CTkFrame(self, width=100, corner_radius=0)  # Reduce width here
        self.bottom_frame.grid(row=5, column=0, columnspan=4, sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        # create a button in the bottom frame
        self.bottom_button = customtkinter.CTkButton(self.bottom_frame, text="Descargar Archivo", width=30, command = self.verif_descargar)
        self.bottom_button.grid(row=0, column=0, padx=10, pady=10, sticky="e")


    def show_checkmark(self):
        CTkMessagebox(message="Archivo Subido Correctamente",
                  icon="check", option_1="Ok", title = "Carga Archivo")
        
    def mensaje_gen(self, message):
        CTkMessagebox(message=message,
                  icon="check", option_1="Ok", title = "Mensaje")
    
    def show_info():
        CTkMessagebox(title="Descarga", message="Archivo Descargado Correctamente")

    def error_archivo(self):
        CTkMessagebox(title="Error", message="Archivo no correcto", icon="cancel")   
    
    def error_gen(self, message):
        CTkMessagebox(title="Error", message=message, icon="cancel")

    def verif_descargar(self):
        self.descargar()
    
        
    def descargar(self):
        try:
            downloads_folder = os.path.expanduser("~\\Downloads")
            filename = "Partidas "+str(self.libro.iloc[2,2])+".xlsx"
            file_path = os.path.join(downloads_folder, filename)
            csv_writer = pd.ExcelWriter(filename, engine = "xlsxwriter")
            self.libro.to_excel(csv_writer, sheet_name = "BI NO. ", startrow=5, index=False, header=True)
            workbook = csv_writer.book
            worksheet= csv_writer.sheets["BI NO. "]
            bold_format = workbook.add_format({'bold': True, 'font_size': 14})
            worksheet.write("B2","TRANSACDIGITAL S.A.", bold_format)
            worksheet.write("B3","Libro de diario", bold_format)
            worksheet.write("B4","(expresado en quetzales)", bold_format)
            
            csv_writer.save()
            
            self.mensaje_gen("Archivo Descargado Correctamente")
        
        except AttributeError:
            self.error_gen("Selecciona los archivos necesarios antes de generar el Excel")
        


    def choose_file1(self):
        filetypes = [("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        self.file1 = file_path
        self.file3 = ""
        self.show_checkmark()
        self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
        self.process_files()


    def choose_file2(self):
        filetypes = [("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        self.file2 = file_path
        self.file3 = ""
        self.show_checkmark()
        self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
        self.process_files()
        
        
    def choose_file3(self):
        filetypes = [("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        self.file3 = file_path
        self.show_checkmark()
        self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
        self.process_files()
        
    
    

    def process_files(self):
        # Check if 3/2 files have been selected
        print(self.file3)
        if (self.file1 != "" and self.file2 != "" and self.file3 != "") and self.tabview.get() == "Cuenta 007":
                
            if self.file1.endswith(".xlsx"):
                # It's an Excel file (.xlsx)
                dfpaggo = pd.read_excel(self.file1)
            elif self.file1.endswith(".csv"):
                # It's a CSV file
                dfpaggo = pd.read_csv(self.file1)
            else:
                # File format not supported
                raise self.error_gen("Unsupported file format")
            
            
            if self.file2.endswith(".xlsx"):
                # It's an Excel file (.xlsx)
                dfbi = pd.read_excel(self.file2)
            elif self.file2.endswith(".csv"):
                # It's a CSV file
                dfbi = pd.read_csv(self.file2)
            else:
                # File format not supported
                raise self.error_gen("Unsupported file format")
                
            if self.file3.endswith(".xlsx"):
                # It's an Excel file (.xlsx)
                dfbitran = pd.read_excel(self.file3, skiprows=range(8), sep=";")
                dfbitran = dfbitran.iloc[:,0:7]
            elif self.file3.endswith(".csv"):
                # It's a CSV file
                dfbitran = pd.read_csv(self.file3, skiprows=range(8), delimiter=";")
                dfbitran = dfbitran.iloc[:,0:7]
            else:
                # File format not supported
                raise self.error_gen("Unsupported file format")
            
            self.check_files(dfbitran, 1, dfpaggo, dfbi)
        
        
        
        elif self.file3 != "" and self.tabview.get() == "Otra":
            if self.file3.endswith(".xlsx"):
                # It's an Excel file (.xlsx)
                dfbitran = pd.read_excel(self.file3, skiprows=range(8), sep=";")
                dfbitran = dfbitran.iloc[:,0:7]
            elif self.file3.endswith(".csv"):
                # It's a CSV file
                dfbitran = pd.read_csv(self.file3, skiprows=range(8), delimiter=";")
                dfbitran = dfbitran.iloc[:,0:7]
            else:
                # File format not supported
                raise self.error_gen("Unsupported file format")
            
            self.check_files(dfbitran, 2)

        else:
            # Clear Treeview when files are not selected
            self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())


    
    def check_files(self, dfbitran, state, dfpaggo="", dfbi=""):
        if state == 1:
            if len(dfpaggo.columns) != 6 and ("Cuenta" not in dfpaggo.columns and "Monto" not in dfpaggo.columns):
                self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
                raise self.error_gen("El archivo de Paggo no contiene la informacion correcta")
            else:
                if len(dfbi.columns) != 12 and ("Cuenta a Acreditar" not in dfbi.columns and "Monto" not in dfbi.columns):
                    self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
                    raise self.error_gen("El archivo de Bi no contiene la informacion correcta")
                else:
                    if len(dfbitran.columns) != 7 and "No. Doc" not in dfbitran.columns:
                        self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
                        raise self.error_gen("El archivo de Bi Transacciones no contiene la informacion correcta")
                    else:
                        self.transform(dfbitran, state, dfpaggo, dfbi)
        elif state == 2:
            if len(dfbitran.columns) != 7 and "No. Doc" not in dfbitran.columns:
                self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
                raise self.error_gen("El archivo de Paggo Transacciones no contiene la informacion correcta")
            else:
                self.transform(dfbitran, state)
        else:
            self.error_gen("Error, Contacte a soporte")



            
    def transform(self, dfbitran, state, dfpaggo="", dfbi=""):
        self.libro = pd.DataFrame(columns=(["Empresa","No.","Mes","partida","No. De partida","Fecha de contabilizacion","Cuenta Contable","Nombre de la cuenta Contable",
                                           "Concepto","Debe (Q.)", "Haber (Q.)", "Tipo de doc","No. de Cheque/Transferencia", "Fecha del documento",
                                           "No. de Serie de Fact.", "No. Factura","Numero de NIT del Proveedor/cliente", "Nombre del Proveedor / Cliente"]))
        espaciado = ["","","","","","","","","","","","","","","","","",""]
        self.libro.loc[len(self.libro)] = espaciado
        self.libro.loc[len(self.libro)] = espaciado
        if state == 1:
            dfbitran.dropna(subset=["Fecha"], inplace=True)
            dfbitran["Fecha"]= pd.to_datetime(dfbitran["Fecha"], format="%d/%m/%Y")
            dfbitran["Mes"] = dfbitran["Fecha"].dt.month
            dfbitran["Fecha"] = dfbitran["Fecha"].dt.date
            dfbitran["Fecha"]= pd.to_datetime(dfbitran["Fecha"])
            dfbitran["Monto"] = dfbitran["Debe (Q.)"].fillna(dfbitran["Haber (Q.)"])
            
            
            df = pd.merge(dfbi,dfpaggo, left_on=["Cuenta a Acreditar","Monto"], right_on=["Cuenta","Monto"])
            df["Fecha proc"]= pd.to_datetime(df["Fecha proc"])
            df["Mes"] = df["Fecha proc"].dt.month
            df["fecha contabilizacion"]=df["Fecha proc"].dt.date
            df["Monto"] = df["Monto"].astype("object")
            df["fecha contabilizacion"] = pd.to_datetime(df["fecha contabilizacion"])
            df["fecha contabilizacion"]=df["Fecha proc"].dt.date
            df["fecha contabilizacion"] = pd.to_datetime(df["fecha contabilizacion"])
            
            #df = pd.merge(df, )
            for row in df.itertuples(index=False):
                partida_debe = ["Transacdigital S.A.","", row[17], "Pd No.", str(row[17])+"-", row[18], "11113", "Cuenta Transitoria", str(row[15])+","+str(row[4])+","+str(row[16]), 
                                row[7], "", "ND", row[2], row[18], "", "", "", ""]
                partida_haber = ["Transacdigital S.A.","", row[17], "Pd No.", str(row[17])+"-", row[18], "11134", "Cuenta Depositos Monetarios BI 0730043007 No. en Q.", str(row[15])+","+str(row[4])+","+str(row[16]),
                                 "", row[7], "NC", row[2], row[18], "", "", "", ""]
                
                self.libro.loc[len(self.libro)] = partida_debe
                self.libro.loc[len(self.libro)] = partida_haber
                self.libro.loc[len(self.libro)] = espaciado
                self.libro.loc[len(self.libro)] = espaciado
                
            self.libro.fillna("", inplace=True)
            self.mensaje_gen("Archivo Procesado y Generado Correctamente")
            self.showdf(self.libro)
        elif state == 2:
            dfbitran.dropna(subset=["Fecha"], inplace=True)
            dfbitran["Fecha"]= pd.to_datetime(dfbitran["Fecha"], format="%d/%m/%Y")
            dfbitran["Mes"] = dfbitran["Fecha"].dt.month
            for row in dfbitran.itertuples(index=False):
                value = str(row[2])
                if "BBWEB" in value:
                    cuenta="11113"
                    nombre = "Cuenta Transitoria"
                else:
                    cuenta = "21115"
                    nombre = "Cuentas Por Pagar PayFAC GT"
                
                
                if pd.isna(row[4]):
                    monto=row[5]
                elif pd.isna(row[5]):
                    monto=row[4]
                
                    
                if row[1] == "NC":
                    debe1=""
                    haber1=row[5]
                    debe2=row[5]
                    haber2=""
                else:
                    debe1=row[4]
                    haber1=""
                    debe2=""
                    haber2=row[4]
                    
                partida_debe = ["Transacdigital S.A.","", row[7], "Pd No.", str(row[7])+"-", row[0], cuenta, nombre, row[2], 
                                debe1, haber1, row[1], row[3], row[0], "", "", "", ""]
                partida_haber = ["Transacdigital S.A.","", row[7], "Pd No.", str(row[7])+"-", row[0], "11134", "Cuenta Depositos Monetarios BI 0730043007 No. en Q.", row[2],
                                 debe2, haber2, row[1], row[3], row[0], "", "", "", ""]
                
                self.libro.loc[len(self.libro)] = partida_debe
                self.libro.loc[len(self.libro)] = partida_haber
                self.libro.loc[len(self.libro)] = espaciado
                self.libro.loc[len(self.libro)] = espaciado
                
            self.libro.fillna("", inplace=True)
            self.mensaje_gen("Archivo Procesado y Generado Correctamente")
            self.showdf(self.libro)
                
        else:
            self.error_gen("Error, contacte a Servicio Tecnico")
    
    
    def showdf(self, df):
        self.dataframe_treeview.delete(*self.dataframe_treeview.get_children())
        
        columns = df.columns.tolist()
        self.dataframe_treeview["columns"] = columns
        for col in columns:
            self.dataframe_treeview.heading(col, text=col)
            self.dataframe_treeview.column(col, width=1)
            
        for row in df.itertuples(index=False):
            self.dataframe_treeview.insert("", tk.END, values=row)
            
        # Update column widths based on content
        for col in columns:
            max_width = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
            self.dataframe_treeview.column(col, width=max_width * 10)  # Adjust the factor to control the width
             
    def destroy(self):
        # Add any cleanup code or termination logic here
        super().destroy()  # Close the main window

if __name__ == "__main__":
    app = App()
    app.mainloop()
