import pandas as pd
import numpy as np
import re
import xlwings as xw
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, Label, Button, messagebox
import os
import sys
import itertools


def add_columns(df):
    global cols_to_add 
    cols_to_add = ['PN','SN1','DES1', 'SN2','DES2', 'SN3','DES3']
    col_index = df.columns.get_loc('Material Description') + 1
    for i, col in enumerate(cols_to_add):
        df.insert(col_index + i, col, '')
    return df


def Fill(level : str, df_mmsta):
    for idx in range(df_mmsta.shape[0]):
        if df_mmsta.loc[:,'Level'][idx] == level:
            # PN_arr.append(df_mmsta.loc[idx]['Material Description'])
            df_mmsta.loc[idx,'PN'] = df_mmsta.loc[idx]['Material Description']
            idx_last_l = idx
        else :
            df_mmsta.loc[idx,'PN'] = df_mmsta.loc[idx_last_l]['Material Description']
            # PN_arr.append(df_mmsta.loc[idx_last_l]['Material Description'])
    return df_mmsta

# def del_vals(mat_desc : str):
#     vals_to_del = ['ctgf', 'gaft', 'cut tube']
#     for v in vals_to_del:
#         if v in mat_desc.lower():
#             return False
#     return True

def Fill_with_col(level : str, cols,df):
    idx_last_l = None
    for idx in range(df.shape[0]):
        if df.loc[idx,'Level'] == level:
            df.loc[idx,cols[0]] = df.loc[idx,'Material']
            df.loc[idx,cols[1]] = df.loc[idx,'Material Description']
            idx_last_l = idx
           
            
        else :
            if idx_last_l is not None:
                df.loc[idx,cols[0]] = df.loc[idx_last_l,'Material']
                df.loc[idx,cols[1]] = df.loc[idx_last_l,'Material Description']
                
            else :
                df.loc[idx,cols[0]] = '(blank)'
                df.loc[idx,cols[1]] = '(blank)'
    return df

def level_sorter(level):
    return int(level[-1])

def separator(MMSTA_path, output_file):
    df_mmsta = pd.read_excel(MMSTA_path)
    df_mmsta = df_mmsta[['Level', 'Material Type' , 'Material', 'Material Description', 'Component quantity']]
    
    df_mmsta = add_columns(df_mmsta)
    df_mmsta = Fill('0',df_mmsta)
    
    # proccessed_log = df_mmsta['Material Description'].apply(del_vals)
    # df_mmsta = df_mmsta[proccessed_log].reset_index(drop=True)

    df_mmsta = df_mmsta.reset_index(drop=True)

    for level in df_mmsta['Level'].unique().tolist():
        sn = 'SN'+ level[-1]
        des = 'DES' + level[-1]
        df_mmsta = Fill_with_col(level, [sn, des], df_mmsta)
        
    df_mmsta = df_mmsta[['Level' ,'PN','SN1','DES1','SN2','DES2','SN3','DES3', 'Component quantity']]

    df_mmsta_pivot = pd.pivot_table(df_mmsta, values = 'Component quantity', index = ['Level','SN1','DES1','SN2' , 'DES2','SN3', 'DES3'], 
                                columns = 'PN', aggfunc = 'size', fill_value = ' ')
    
    wb = xw.Book()
    sheet_mmsta = wb.sheets[0]
    sheet_mmsta.name = "sep mmsta"
    
    df_mmsta_pivot = df_mmsta_pivot.reset_index()
    
    df_mmsta_pivot.loc[df_mmsta_pivot['Level'] == '*1', ['SN2', 'DES2', 'SN3', 'DES3']] = '(blank)'
    df_mmsta_pivot.loc[df_mmsta_pivot['Level'] == '**2', ['SN3', 'DES3']] = '(blank)'
    
    
    df_mmsta_pivot['level_sort'] = df_mmsta_pivot['Level'].apply(level_sorter)
    df_mmsta_pivot = df_mmsta_pivot.sort_values('level_sort').drop(columns='level_sort')
    
    df_mmsta_pivot[df_mmsta_pivot.columns[7:]] = df_mmsta_pivot[df_mmsta_pivot.columns[7:]].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    
    columns_to_agg = list(df_mmsta_pivot.columns[7:])
    cols = ['Level', 'SN1', 'DES1', 'SN2', 'DES2', 'SN3', 'DES3']
    cols.extend(columns_to_agg)
    
    agg_dict = {col: 'sum' for col in columns_to_agg}

    df_mmsta_pivot = df_mmsta_pivot.groupby(['Level','SN1','SN2','SN3']).agg({
        'DES1': 'first',
        'DES2': 'first',
        'DES3': 'first',
        **agg_dict
        
    }).reset_index()

    df_mmsta_pivot = df_mmsta_pivot[cols]
    df_mmsta_pivot.replace(to_replace=" " , value="(Blank)")
    df_mmsta_pivot.replace(to_replace= 0 , value=" ")


    sheet_mmsta.range("A1").value = [df_mmsta_pivot.columns.tolist()] + df_mmsta_pivot.values.tolist()
    
    df_fil_simple = df_mmsta_pivot [df_mmsta_pivot['DES1'].apply(lambda x : 'circuit' in x.lower())]
    columns_to_agg = list(df_fil_simple.columns[7:])

    agg_dict = {col: 'sum' for col in columns_to_agg}

    df_fil_simple = df_fil_simple.groupby(['Level','SN1']).agg({
        'DES1': 'first',
        'SN2': 'first',
        'DES2': 'first',
        'SN3': 'first',
        'DES3': 'first',
        **agg_dict
        
    }).reset_index()

    df_fil_simple.replace(to_replace=" " , value="(Blank)")
    df_fil_simple.replace(to_replace= 0 , value=" ")

    double_sheet = wb.sheets.add(after=sheet_mmsta)
    sheet_mmsta.name = "Fil Simple"
    sheet_mmsta.range("A1").value = [df_fil_simple.columns.tolist()] + df_fil_simple.values.tolist()
    
    df_double = df_mmsta_pivot [df_mmsta_pivot['DES1'].apply(lambda x : 'double' in x.lower())]
    df_double = df_double.groupby(['Level','SN1']).agg({
        'DES1': 'first',
        'SN2': 'first',
        'DES2': 'first',
        'SN3': 'first',
        'DES3': 'first',
        **agg_dict
        
    }).reset_index()

    double_sheet = wb.sheets.add(after=sheet_mmsta)
    double_sheet.name = "Double"
    double_sheet.range("A1").value = [df_double.columns.tolist()] + df_double.values.tolist()
    
    df_twisted = df_mmsta_pivot [df_mmsta_pivot['DES1'].apply(lambda x : 'twisted' in x.lower())]
    df_twisted = df_twisted.groupby(['Level','SN1','SN2','SN3']).agg({
        'DES1': 'first',
        'DES2': 'first',
        'DES3': 'first',
        **agg_dict
        
    }).reset_index()
    twisted_sheet = wb.sheets.add(after=double_sheet)
    twisted_sheet.name = "Twisted"
    twisted_sheet.range("A1").value = [df_twisted.columns.tolist()] + df_twisted.values.tolist()
    
    df_joint = df_mmsta_pivot [df_mmsta_pivot['DES1'].apply(lambda x : 'joint' in x.lower())]
    df_joint = df_joint.groupby(['Level','SN1','SN2','SN3']).agg({
        'DES1': 'first',
        'DES2': 'first',
        'DES3': 'first',
        **agg_dict
        
    }).reset_index()
    joint_sheet = wb.sheets.add(after=twisted_sheet)
    joint_sheet.name = "joint"
    joint_sheet.range("A1").value = [df_joint.columns.tolist()] + df_joint.values.tolist()
    
    df_super_grp = df_mmsta_pivot [df_mmsta_pivot['DES1'].apply(lambda x : 'super group' in x.lower())]
    df_super_grp = df_super_grp.groupby(['Level','SN1','SN2','SN3']).agg({
        'DES1': 'first',
        'DES2': 'first',
        'DES3': 'first',
        **agg_dict
        
    }).reset_index()
    superGrp_sheet = wb.sheets.add(after=joint_sheet)
    superGrp_sheet.name = "Super group"
    superGrp_sheet.range("A1").value = [df_super_grp.columns.tolist()] + df_super_grp.values.tolist()
    wb.save(output_file)
    wb.close()

def resource_path(relative_path):
    try:
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    except Exception as e:
        print(f"Erreur lors du chargement de la ressource : {e}")
        return relative_path
    
def select_BOM_file():
    global BOM_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
    
    if not file_path.lower().endswith((".xlsx", ".xls")):
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide (*.xlsx, *.xls;*.XLSX)")
        return

    BOM_file_path = file_path
    ref_label.config(text=f"MMSTA File: {BOM_file_path}")    
    
def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory()
    output_label.config(text=f"Output Folder : {output_dir}")

def main_separator():
    if not BOM_file_path or not output_dir:
        messagebox.showerror("Erreur", "Veuillez sélectionner le fichier MMSTA et le dossier de sortie")
        return
    status_text = f"MMSTA file: {BOM_file_path}\n"   
    if output_dir:
        status_text += f"\nOutput folder: {output_dir}" 

    # status_label.config(text=status_text, fg="black")
    
    try:
        output_file = f"{output_dir}/output.xlsx"
        print("Intégration des fichiers...")
        print(BOM_file_path)
        print(output_file)
        messagebox.showinfo("showinfo", "Intégration des fichiers en cours...")
        separator(BOM_file_path,output_file)
        messagebox.showinfo("Succès", "Intégration des fichiers terminée avec succès !")
        return output_file
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'intégration des fichiers: {e}")
        return
    
def Read_wire_list(path_wire_list):
    #df_wire_list = pd.read_excel('../Data/2D0062156P-H0000_MaxWireList.xlsx')
    print(path_wire_list)
    df_wire_list = pd.read_excel(path_wire_list)
    PN_columns_wirelist = df_wire_list.columns[54:]
    PN_columns_wirelist = list(PN_columns_wirelist)
    PN_columns_wirelist = [col.split(':')[0] for col in PN_columns_wirelist]

    num_cols_to_rename = len(df_wire_list.columns[54:])
    df_wire_list.columns = df_wire_list.columns[:54].tolist() + PN_columns_wirelist

    loc = df_wire_list.columns.get_loc('Wire Internal Name')
    df_wire_list.insert(loc + 1, 'DES', None, allow_duplicates=False)
    df_wire_list.insert(loc + 1, 'Sup Grp', None, allow_duplicates=False)
    df_wire_list.insert(loc + 1, 'SN2', None, allow_duplicates=False)
    df_wire_list.insert(loc + 1, 'SN1', None, allow_duplicates=False)

    df_wire_list['Wire Internal Name'] = df_wire_list['Wire Internal Name'].str.replace('W', '').astype(int)

    df_wire_list[PN_columns_wirelist] = df_wire_list[PN_columns_wirelist].replace('X', 1).fillna(0).astype(int)
    df_wire_list[PN_columns_wirelist] = df_wire_list[PN_columns_wirelist].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    return df_wire_list

def get_col(wire,material_to_comp_dict,mmsta_sep):
     cols = []
     keys_for_Wire = [key for key, value in material_to_comp_dict.items() if value == wire]
     for key in keys_for_Wire :
          col = mmsta_sep.columns[mmsta_sep.isin([key]).any()][0]
          cols.append(col)

     return cols

# Define a function to find rows in df_wire_list that match product number patterns in mmsta_sep
def find_matching_rows_for_wire(wire_internal_name, df_wire_list, mmsta_sep, PN_columns_wirelist, PN_columns_mmsta):

    matches = []
    
    # Get rows for the specified wire
    wire_rows = df_wire_list[df_wire_list['Wire Internal Name'] == wire_internal_name]
    matching_rows = mmsta_sep[mmsta_sep['Wire Internal Name'] == wire_internal_name]
    
    if not matching_rows.empty and not wire_rows.empty:
        for wire_idx in wire_rows.index:
            pene_max = wire_rows.loc[wire_idx, PN_columns_wirelist].values
            
            for mmsta_idx in matching_rows.index:
                pen_mmsta = matching_rows.loc[mmsta_idx, PN_columns_mmsta].values
                
                # Check if product number patterns match
                if np.array_equal(pene_max, pen_mmsta):
                    matches.append((wire_idx, mmsta_idx))
    
    return matches

def processing_function(mmsta_sep, df_wire_list, PN_columns_wirelist, PN_columns_mmsta):
      all_wires = []

      for wire in df_wire_list['Wire Internal Name'].unique():
            df_wire = df_wire_list[df_wire_list['Wire Internal Name'] == wire]
            matching_rows = mmsta_sep[mmsta_sep['Wire Internal Name'] == wire]
            if df_wire.shape[0] == 1 and not matching_rows.empty:
                  pene_max = df_wire[PN_columns_wirelist].values[0]
                  pen_mmsta = np.sum(matching_rows[PN_columns_mmsta].values, axis=0)
                  if np.array_equal(pene_max, pen_mmsta):
                        n_fois_duplicate = mmsta_sep[mmsta_sep['Wire Internal Name'] == wire].shape[0]

                        if n_fois_duplicate:
                  
                              row_to_duplicate = df_wire.copy()
                              df_wire_list = df_wire_list[df_wire_list['Wire Internal Name'] != wire]
                              # df_wire_list = df_wire_list.reset_index(drop=True)
                        
                              for i in range(n_fois_duplicate):
                                    matching_row = matching_rows.iloc[i]
                                    level = matching_row['Level'][-1]
                                    row_to_duplicate[PN_columns_wirelist] = matching_row[PN_columns_mmsta].values
                                    if int(level) == 3:
                                          row_to_duplicate['SN1'] = matching_row['SN3'] 
                                          row_to_duplicate['SN2'] = matching_row['SN2']
                                          row_to_duplicate['Sup Grp'] = matching_row['SN1']
                                          row_to_duplicate['DES'] = matching_row['DES3']
                                    elif int(level) == 2:
                                          row_to_duplicate['SN1'] = matching_row['SN2']
                                          row_to_duplicate['SN2'] = matching_row['SN1']
                                          row_to_duplicate['DES'] = matching_row['DES2']
                                    elif int(level) == 1:
                                          row_to_duplicate['SN1'] = matching_row['SN1']
                                          row_to_duplicate['DES'] = matching_row['DES1']

                              
                                    df_wire_list = pd.concat([df_wire_list, row_to_duplicate])
                                    df_wire_list = df_wire_list.infer_objects()

                             
            elif df_wire.shape[0] > 1 and not matching_rows.empty:
                  if matching_rows.shape[0] >= df_wire.shape[0]: #matching_rows.shape[0] > df_wire.shape[0]:
                        indices_valides = get_idx(df_wire,matching_rows, PN_columns_wirelist, PN_columns_mmsta) #(idx_max_wire,(les indices de mmsta))
                        if indices_valides :
                              for idx_val in indices_valides :
                                    n_fois_duplicate = len(idx_val[1])
                                    idx_mmsta = idx_val[1]
                                    idx_row_to_duplicate = idx_val[0]
                                    if n_fois_duplicate >= 1:
                                          row_to_duplicate = df_wire.loc[idx_row_to_duplicate]
                                          df_wire_list = df_wire_list.drop(idx_row_to_duplicate)
                                          # df_wire_list = df_wire_list.reset_index(drop=True)
                                    
                                          for i in list(idx_mmsta):
                                                matching_row = matching_rows.loc[i]
                                                level = matching_row['Level'][-1]
                                                row_to_duplicate[PN_columns_wirelist] = matching_row[PN_columns_mmsta].values
                                                if int(level) == 3:
                                                      row_to_duplicate['SN1'] = matching_row['SN3'] 
                                                      row_to_duplicate['SN2'] = matching_row['SN2']
                                                      row_to_duplicate['Sup Grp'] = matching_row['SN1']
                                                      row_to_duplicate['DES'] = matching_row['DES3']
                                                elif int(level) == 2:
                                                      row_to_duplicate['SN1'] = matching_row['SN2']
                                                      row_to_duplicate['SN2'] = matching_row['SN1']
                                                      row_to_duplicate['DES'] = matching_row['DES2']
                                                elif int(level) == 1:
                                                      row_to_duplicate['SN1'] = matching_row['SN1']
                                                      row_to_duplicate['DES'] = matching_row['DES1']
                                          
                                                df_wire_list = pd.concat([df_wire_list, row_to_duplicate.to_frame().T])
                                                df_wire_list = df_wire_list.infer_objects()
                                                # df_wire_list = df_wire_list.reset_index(drop=True)
                        
                        
                  # else:
                  #       tuple_match = find_matching_rows_for_wire(wire, df_wire_list, mmsta_sep, PN_columns_wirelist, PN_columns_mmsta)
                  #       for wire_idx, mmsta_idx in tuple_match:
                  #             matching_row = mmsta_sep.loc[mmsta_idx]
                  #             # 127 __ 956 __ 151
                  #             # 128 __ 574 __ 151
                  #             # # Create a new row based on the matching row
                  #             new_row = df_wire_list.loc[wire_idx]
                  #             df_wire_list = df_wire_list.drop(wire_idx)#df_wire_list[df_wire_list['Wire Internal Name'] != wire]
                  #             # df_wire_list = df_wire_list.reset_index(drop=True)
                  #             new_row[PN_columns_wirelist] = matching_row[PN_columns_mmsta].values
                              
                  #             level = matching_row['Level'][-1]
                              
                  #             if int(level) == 3:
                  #                   new_row['SN1'] = matching_row['SN3'] 
                  #                   new_row['SN2'] = matching_row['SN2']
                  #                   new_row['Sup Grp'] = matching_row['SN1']
                  #                   new_row['DES'] = matching_row['DES3']
                  #                   new_row['SN1'] = matching_row['SN2']
                  #                   new_row['SN2'] = matching_row['SN1']
                  #                   new_row['DES'] = matching_row['DES2']
                  #             elif int(level) == 1:
                  #                   new_row['SN1'] = matching_row['SN1']
                  #                   new_row['DES'] = matching_row['DES1']
                              
                  #             # Append the new row to df_wire_list
                  #             df_wire_list = pd.concat([df_wire_list, new_row.to_frame().T])
                  #             # df_wire_list = df_wire_list.reset_index(drop=True)
      return df_wire_list
            

def get_idx(df_wire,matching_rows,PN_columns_wirelist,PN_columns_mmsta):
    indices_valides = []
    for idx in df_wire.index:
        target = df_wire.loc[idx, PN_columns_wirelist].values
        for r in range(1, len(matching_rows) + 1):  
            for indices in itertools.combinations(matching_rows.index, r):
                pene = matching_rows.loc[list(indices)][PN_columns_mmsta].values
                sum_pene = np.sum(pene, axis=0)
                if np.array_equal(sum_pene, target):
                    indices_valides.append((idx,indices))

    #print("Indices des lignes valides :", indices_valides)
    return indices_valides

def integrator(path_mmsta_sep, path_wire_list, path_ypp_cae, Output_integrated):
    # df_ypp_cae = pd.read_excel('../Data/YPP_CAE 1 (2).XLSX', sheet_name='Sheet1')
    df_ypp_cae = pd.read_excel(path_ypp_cae, sheet_name='Sheet1')
    df_ypp_cae_pivot = pd.pivot_table(df_ypp_cae, values = 'Counter',index = ['Material','Composition No'],
                                columns = 'Product number', aggfunc = 'count', fill_value = ' ')
    
    df_ypp_cae_pivot.reset_index(inplace=True)
    
    df_ypp_cae_pivot[df_ypp_cae_pivot.columns[2:]] = df_ypp_cae_pivot[df_ypp_cae_pivot.columns[2:]].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    
    PN_columns = df_ypp_cae_pivot.columns[2:]
    PN_columns = list(PN_columns)
    
    material_to_comp_dict = dict(zip(df_ypp_cae_pivot['Material'], df_ypp_cae_pivot['Composition No']))

    # Read MMSTA
    #mmsta_sep = pd.read_excel('../Data/mmsta_sep.xlsx')
    mmsta_sep = pd.read_excel(path_mmsta_sep)
    mmsta_sep['Wire Internal Name'] = mmsta_sep['SN1'].map(material_to_comp_dict)

    idx_1 = mmsta_sep[mmsta_sep['Wire Internal Name'].isnull()].index
    mmsta_sep.loc[idx_1, 'Wire Internal Name'] = mmsta_sep.loc[idx_1, 'SN2'].map(material_to_comp_dict)
    
    idx_2 = mmsta_sep[mmsta_sep['Wire Internal Name'].isnull()].index
    mmsta_sep.loc[idx_2, 'Wire Internal Name'] = mmsta_sep.loc[idx_2, 'SN3'].map(material_to_comp_dict)
    
    PN_columns_mmsta = mmsta_sep.columns[7:-1]
    PN_columns_mmsta = list(PN_columns_mmsta)
    
    df_wire_list = Read_wire_list(path_wire_list)
    
    df_wire_list = processing_function(mmsta_sep, df_wire_list, PN_columns, PN_columns_mmsta)
    
    df_wire_list = df_wire_list.reset_index(drop=True)
    #df_wire_list.to_excel('../Data/Updated_2.xlsx', index=False)     
    df_wire_list.to_excel(Output_integrated, index=False)     
    

def select_ypp_cae_file():
    
    global ypp_cae_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
    if not file_path.lower().endswith((".xlsx", ".xls")):
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide (*.xlsx, *.xls;*.XLSX)")
        return

    ypp_cae_file_path = file_path
    ypp_label.config(text=f"YPP CAE File: {os.path.basename(file_path)}")
    
    
def select_mmsta_sep_file():
    global mmsta_sep_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
    if not file_path.lower().endswith((".xlsx", ".xls")):
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide (*.xlsx, *.xls;*.XLSX)")
        return

    mmsta_sep_file_path = file_path
    mmsta_sep_label.config(text=f"MMSTA Sep File: {os.path.basename(file_path)}")
    
def select_wire_list_file():
    global wire_list_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
    if not file_path.lower().endswith((".xlsx", ".xls")):
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide (*.xlsx, *.xls;*.XLSX)")
        return

    wire_list_file_path = file_path
    wire_list_label.config(text=f"MMSTA Sep File: {os.path.basename(file_path)}")   

def main_integrator():
    if not all([ypp_cae_file_path, mmsta_sep_file_path, wire_list_file_path, output_dir]):
        messagebox.showerror("Error", "Please select all required files and output folder")
        return

    try:
        output_file = os.path.join(output_dir, "integrated_output.xlsx")
        integrator(
                mmsta_sep_file_path, 
                wire_list_file_path, 
                ypp_cae_file_path, 
                output_file
            )
        messagebox.showinfo("Success", f"Files integrated successfully!\nOutput: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error during integration: {str(e)}")


logo_path = resource_path("../assets/yazaki_logo.png")

root = tk.Tk()
root.title("Data Integration - Yazaki")
root.geometry("350x600")

try:
    img = Image.open(logo_path)
    img = img.resize((150, 50), Image.LANCZOS)
    logo = ImageTk.PhotoImage(img)
    root.logo = logo
    logo_label = Label(root, image=logo)
    logo_label.pack(pady=10)
except:
    logo_label = Label(root, text="[Logo Yazaki]", font=("Arial", 14, "bold"))
    logo_label.pack(pady=10)


wire_list_file_path = ""
mmsta_sep_file_path = ""
ypp_cae_file_path = ""
BOM_file_path = ""
output_dir = ""

ref_button = Button(root, text="Sélectionner le fichier MMSTA", command=select_BOM_file)
ref_button.pack(pady=5)
ref_label = Label(root, text="Référence: Non sélectionné", wraplength=400)
ref_label.pack()

output_button = Button(root, text="Sélectionner le dossier de sortie", command=select_output_dir)
output_button.pack(pady=5)
output_label = Label(root, text="Dossier de sortie: Non sélectionné", wraplength=400)
output_label.pack()

compare_button = Button(root, text="Separate", command=main_separator)
compare_button.pack(pady=20)

separator_frame = tk.Frame(root)
separator_frame.pack(padx=20, pady=10)

Label(separator_frame, text="MMSTA Separator", font=("Arial", 12, "bold")).pack(pady=10)
Button(separator_frame, text="Select YPP CAE File", command=select_ypp_cae_file).pack(pady=5)
ypp_label = Label(separator_frame, text="YPP CAE File: Not selected", wraplength=350)
ypp_label.pack()
        
Button(separator_frame, text="Select MMSTA Sep File", command=select_mmsta_sep_file).pack(pady=5)
mmsta_sep_label = Label(separator_frame, text="MMSTA Sep File: Not selected", wraplength=350)
mmsta_sep_label.pack()

integrator_frame = tk.Frame(root)
integrator_frame.pack(padx=20, pady=10)
Button(integrator_frame, text="Select Wire List File", command=select_wire_list_file).pack(pady=5)
wire_list_label = Label(integrator_frame, text="Wire List File: Not selected", wraplength=350)
wire_list_label.pack()
        
Button(integrator_frame, text="Integrate Files", command=main_integrator).pack(pady=10)

status_label = Label(root, text="", fg="black")
status_label.pack()

tk.mainloop()

# MMSTA_path = "../Data/MMSTA LOWDASH RHN DORA.XLSX"
# 
    
    
    