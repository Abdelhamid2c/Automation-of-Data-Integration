import xlwings as xw
import os
import tempfile
import re


def color_connector_cavities_corrected(ws):

    print(f"Traitement de la feuille: {ws.name}")
    
    # --- Phase de nettoyage ---
    for shape in ws.shapes:
        try:
            if shape.name.startswith("DiagLine_Cavity_"):
                shape.delete()
        except:
            pass
    
    # --- Phase de traitement ---
    start_row = 16
    current_row = start_row
    
    while ws.range(f"A{current_row}").value is not None and str(ws.range(f"A{current_row}").value).strip() != "":
        try:
            cavity_number = ws.range(f"A{current_row}").value
            cell_b = ws.range(f"B{current_row}")
            
            # Obtenir la couleur de fond
            background_color = cell_b.color
            
            
            if background_color == (0, 0, 0) : 
                print("Noir")
                bg_color_rgb = 0
            elif background_color == (255, 255, 255) :
                bg_color_rgb = 16777215
            else:
                # Conversion standard RGB pour Excel
                r, g, b = background_color
                bg_color_rgb = r + (g << 8) + (b << 16)
            
            # Vérifier si une bordure diagonale montante existe
            has_diagonal = False
            try:
                # 6 = xlDiagonalUp (diagonale montante)
                border_style = cell_b.api.Borders(6).LineStyle
                
                if border_style != -4142:  # -4142 = xlNone
                    line_color = cell_b.api.Borders(6).Color
                    has_diagonal = True
                    print("Diagonal from bottom-left to top-right")
                else:
                    # Si pas de diagonale, ne pas définir de couleur de ligne
                    has_diagonal = False
            except:
                has_diagonal = False
            
            
            if not has_diagonal:
                line_color = 0
                border_style = cell_b.api.Borders(5).LineStyle
                
                if border_style != -4142:  # -4142 = xlNone
                    line_color = cell_b.api.Borders(5).Color
                    has_diagonal = True
                    print("Diagonal from top-left to bottom-right")
                else:
                    # Si pas de diagonale, ne pas définir de couleur de ligne
                    has_diagonal = False
                
            
            found_shape = False
            for shape in ws.shapes:
                try:
                    if shape.api.Type == 6:  
                        continue
                    shape_digits = re.sub(r'[^0-9]', '', shape.name)
                    if shape_digits:
                        shape_digit_str = str(shape_digits)
                        print(shape_digit_str)
                    
                    # if hasattr(shape.api, "TextFrame2") and shape.api.TextFrame2.HasText:
                    #     shape_text = shape.api.TextFrame2.TextRange.Text
                        if shape_digit_str.strip() == str(int(cavity_number)).strip():
                            found_shape = True
                            
                            print(f"Trouvé: Cavité {str(int(cavity_number)).strip()} correspond à forme avec texte '{shape_digit_str.strip()}'")
                            shape.api.Fill.Visible = True
                            shape.api.Fill.Solid()
                            shape.api.Fill.ForeColor.RGB = bg_color_rgb
                            print(f"Couleur de fond: {bg_color_rgb}")
                            
                            # Masquer le contour de la forme
                            shape.api.Line.Visible = False
                            
                            # Dessiner la diagonale seulement si elle existe
                            if has_diagonal:
                                x1 = shape.api.Left
                                y1 = shape.api.Top + shape.api.Height
                                x2 = shape.api.Left + shape.api.Width
                                y2 = shape.api.Top
                                
                                # Ajouter la ligne
                                diag_line = ws.shapes.api.AddLine(x1, y1, x2, y2)
                                
                                # Formater la ligne
                                diag_line.Name = f"DiagLine_Cavity_{cavity_number}"
                                diag_line.Line.ForeColor.RGB = line_color
                                diag_line.Line.Weight = 2.5
                                diag_line.Line.Visible = True
                            
                            break
                except Exception as e:
                    print(f"Erreur avec forme: {e}")
                    continue
            
            if not found_shape:
                print(f"Attention: Forme pour Cavité '{cavity_number}' non trouvée.")
        except Exception as e:
            print(f"Erreur en ligne {current_row}: {e}")
        
        current_row += 1
    
    print("Coloration terminée.")
    return ws

def color_all_Cavities(path):
    wb = xw.Book(path)
    ws = wb.sheets
    for sheet in ws:
        ws = wb.sheets[sheet.name]  
        print(ws)
        color_connector_cavities_corrected(ws)
    wb.save(path)
    
    
def find_first_empty_row(wb, sheet_name=None):
    try:
        
        if sheet_name:
            sheet = wb.sheets[sheet_name]
        else:
            sheet = wb.sheets.active
        
        last_row = sheet.used_range.last_cell.row
        
        image_rows = set()
        if sheet.shapes:
            for pic in sheet.shapes:
                top_row = 1
                bottom_row = 1
                
                while sheet.range(f"A{top_row}").top < pic.top:
                    top_row += 1
                
                bottom_row = top_row
                pic_bottom = pic.top + pic.height
                while bottom_row >= last_row and sheet.range(f"A{bottom_row}").top < pic_bottom:
                    bottom_row += 1
                
                for r in range(top_row-1, bottom_row+1):
                    image_rows.add(r)
            print(image_rows)
        
        else :
            # for row in range(1, last_row + 2):
            #     if (sheet.range(f"A{row}").value is None) and (row not in image_rows):
            #         return row
            return last_row
        
        return max(last_row + 1, max(image_rows) - 1 if image_rows else 0)
        
    except Exception as e:
        print(f"Error: {e}")
        return None


# output_file = r"C:\Users\user\Desktop\Connecters\output.xlsx"
# output_sheet_name = "sps"

# first_empty_row = find_first_empty_row(output_file, output_sheet_name)
# print(f"The first empty row in column A is: A{first_empty_row}")


def capture_sheet_as_image(workbook_path, source_sheet_name, target_path,target_sheet_name,position):

    workbook_path = os.path.abspath(workbook_path)

    workbook_path_dest = os.path.abspath(target_path)    
    wb_target = xw.Book(workbook_path_dest)
    # Verify the workbook exists
    if not os.path.exists(workbook_path):
        print(f"Error: Workbook not found at '{workbook_path}'")
        return False
    
    try:
        # Open the workbook
        wb = xw.Book(workbook_path)
        print(f"Opened workbook: {wb.name}")
        
        # Get source and target sheets
        try:
            source_sheet = wb.sheets[source_sheet_name]
            target_sheet = wb_target.sheets[target_sheet_name]
        except Exception as e:
            print(f"Error finding sheets: {e}")
            return False
        
        # Activate source sheet and get the used range
        source_sheet.activate()
        used_range = source_sheet.used_range
        
        # Method 1: Using clipboard (more reliable)
        try:
            # Copy as picture to clipboard
            used_range.api.CopyPicture()
            
            target_sheet.activate()
            target_sheet.range(position).api.Select()
            target_sheet.api.Paste()
            
            print(f"Successfully copied image from '{source_sheet_name}' to '{target_sheet_name}'")
            wb.save()
            return True
            
        except Exception as e:
            print(f"Error with clipboard method: {e}")
            
            # Method 2: Using temporary file as fallback
            try:
                # Create a temporary file path in the system temp directory
                temp_dir = tempfile.gettempdir()
                temp_image_path = os.path.join(temp_dir, "excel_temp_image.png")
                
                print(f"Trying to save image to: {temp_image_path}")
                
                # Export to image
                source_sheet.api.Export(temp_image_path)
                position = target_sheet.range(position).api.TopLeftCell
                target_sheet.pictures.add(
                    temp_image_path,
                    name='ScreenshotImage',
                    left=0,
                    top= target_sheet.range(f"A{position}").top,
                )
                
                try:
                    os.remove(temp_image_path)
                except:
                    pass
                
                print(f"Successfully captured image from '{source_sheet_name}' to '{target_sheet_name}'")
                wb.save()
                return True
                
            except Exception as e2:
                print(f"Error with temporary file method: {e2}")
                return False
    
    except Exception as e:
        print(f"Unexpected error: {e}")
        return False


def get_sheet_names(wb):
    try:
        sheet_names = [sheet.name for sheet in wb.sheets]
        return sheet_names
        
    except Exception as e:
        print(f"Error getting sheet names: {e}")
        return []
    
    

def get_connecteurs(liste_connecteurs, chemin_connecteurs, output_file, output_sheet_name):
    chemin_connecteurs = os.path.abspath(chemin_connecteurs)
    output_file = os.path.abspath(output_file)
    
    if not os.path.exists(chemin_connecteurs):
        print(f"Erreur: Fichier connecteurs non trouvé: '{chemin_connecteurs}'")
        return False
        
    if not os.path.exists(output_file):
        print(f"Erreur: Fichier de sortie non trouvé: '{output_file}'")
        return False
    
    
    try :
        wb_connecteurs = xw.Book(chemin_connecteurs)
        wb_output = xw.Book(output_file)
        print(f"Fichiers ouverts: {wb_connecteurs.name} et {wb_output.name}")
        print(get_sheet_names(wb_connecteurs))
        
        for connecteur in liste_connecteurs:
            source_sheet_name = connecteur
            
            if source_sheet_name not in get_sheet_names(wb_connecteurs):
                print(f"Erreur: Feuille '{source_sheet_name}' non trouvée dans '{chemin_connecteurs}'")
                continue
            first_empty_row = find_first_empty_row(wb_output, output_sheet_name)
            find_first_empty_row(wb_output, output_sheet_name)
            position = f"A{first_empty_row}"
            
            capture_sheet_as_image(chemin_connecteurs, source_sheet_name, output_file, output_sheet_name,position)
            print(f"Image de '{source_sheet_name}' capturée dans '{output_sheet_name}' à la position {position}")
            wb_output.save()
        
    except Exception as e:
        print(f"Erreur: {e}")
        return False

# liste_connecteurs = ["C72","C140", "C46","C3333"]
# chemin_connecteurs = r"C:\Users\user\Desktop\Connecters\Copy_c2.xlsx"
# output_file = r"C:\Users\user\Desktop\Connecters\output.xlsx"
# output_sheet_name = "sps"
# get_connecteurs(liste_connecteurs, chemin_connecteurs, output_file, output_sheet_name)


def remove_diagonal_line_for_cavity(ws, cavity_number):

    line_name = f"DiagLine_Cavity_{cavity_number}"
    
    for shape in ws.shapes:
        try:
            if shape.name == line_name:
                print(f"Found diagonal line for cavity {cavity_number}, removing it")
                shape.delete()
                return True
        except Exception as e:
            print(f"Error processing shape: {e}")
    
    return False

def remove_fill_from_all_shapes(ws):

    print(f"Processing worksheet: {ws.name}")
    
    for shape in ws.shapes:
        digits = re.match(r'^\d+$', shape.name)
        if digits:
            digits = digits.group(0)
            shape.api.Fill.Visible = False
            remove_diagonal_line_for_cavity(ws, float(shape.name))
            print(f"Shape '{shape.name}': {digits}") 
    else :
        print(f"Shape '{shape.name}': No digits found")
       


# path = r"C:\Users\user\Desktop\Connecters\Copy_c2.xlsx"
# wb = xw.Book(path)
# ws = wb.sheets["C46"] 