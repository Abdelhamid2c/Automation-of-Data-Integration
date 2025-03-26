def Fill_with_col_optimized(level: str, cols, df):
    level_mask = df['Level'] == level
    
    df.loc[level_mask, cols[0]] = df.loc[level_mask, 'Material']
    df.loc[level_mask, cols[1]] = df.loc[level_mask, 'Material Description']
    last_level_idx = pd.Series(index=df.index, dtype='int64')
    
    last_level_idx.iloc[0] = -1 if not level_mask.iloc[0] else 0
    
    for i in range(1, len(df)):
        if level_mask.iloc[i]:
            last_level_idx.iloc[i] = i
        else:
            last_level_idx.iloc[i] = last_level_idx.iloc[i-1]
    
    non_level_with_prev = (~level_mask) & (last_level_idx >= 0)
    
    for i in df.index[non_level_with_prev]:
        prev_idx = last_level_idx.loc[i]
        if prev_idx >= 0:
            df.loc[i, cols[0]] = df.loc[prev_idx, 'Material']
            df.loc[i, cols[1]] = df.loc[prev_idx, 'Material Description']
    
    blank_mask = (~level_mask) & (last_level_idx < 0)
    df.loc[blank_mask, cols[0]] = '(blank)'
    df.loc[blank_mask, cols[1]] = '(blank)'
    
    return df




# Compare the two dataframes df_mmsta and df_mmsta_t to check if they are equal
are_dfs_equal = df_mmsta.equals(df_mmsta_t)
print(f"Are df_mmsta and df_mmsta_t equal? {are_dfs_equal}")

# For a more detailed comparison, check which columns have differences
if not are_dfs_equal:
    # Compare each column
    for col in df_mmsta.columns:
        col_equal = df_mmsta[col].equals(df_mmsta_t[col])
        print(f"Column '{col}' is equal: {col_equal}")
    
    # Sample of rows where differences exist
    diff_mask = (df_mmsta != df_mmsta_t).any(axis=1)
    if diff_mask.any():
        print("\nSample of rows with differences:")
        print(df_mmsta[diff_mask].head())
        print("\nCorresponding rows in df_mmsta_t:")
        print(df_mmsta_t[diff_mask].head())
        
        
        

wire_rows = df_wire_list[df_wire_list['Wire Internal Name'] == 347]
    # display(wire_rows)
matching_rows = mmsta_sep[mmsta_sep['Wire Internal Name'] == 347]

def get_idx(df_wire):
    indices_valides = []
    for idx in df_wire.index:
        target = df_wire.loc[idx, PN_columns_wirelist].values
        for r in range(1, len(matching_rows) + 1):  
            for indices in itertools.combinations(matching_rows.index, r):
                pene = matching_rows.loc[list(indices)][PN_columns_mmsta].values
                sum_pene = np.sum(pene, axis=0)
                if np.array_equal(sum_pene, target):
                    indices_valides.append((idx,indices))

    print("Indices des lignes valides :", indices_valides)
    return indices_valides



def get_idx():
    indices_valides = []
    for idx in df_wire.index:
        target = df_wire.loc[idx, PN_columns_wirelist].values
        for r in range(1, len(matching_rows) + 1):  
            for indices in itertools.combinations(matching_rows.index, r):
                pene = matching_rows.loc[list(indices)][PN_columns_mmsta].values
                sum_pene = np.sum(pene, axis=0)
                if np.array_equal(sum_pene, target):
                    indices_valides.append((idx,indices))

    print("Indices des lignes valides :", indices_valides)
    return indices_valides




# Fix the code in cell 34 which has syntax errors and logical issues
import numpy as np
all_wires = []
data_integrated = pd.DataFrame(columns=df_wire_list.columns)

for wire in df_wire_list['Wire Internal Name'].unique():
      df_wire = df_wire_list[df_wire_list['Wire Internal Name'] == wire]
      matching_rows = mmsta_sep[mmsta_sep['Wire Internal Name'] == wire]
      # if wire == 151:
      #       all_wires.append(wire)
      #       print(f'shape of dfw_ire = {df_wire.shape[0]} __ {matching_rows.shape[0]}')
      #       print(matching_rows)
      #       break
      # else :
      #       continue
      if df_wire.shape[0] == 1 and not matching_rows.empty:
            pene_max = df_wire_list[df_wire_list['Wire Internal Name'] == wire][PN_columns_wirelist].values[0]
            pen_mmsta = np.sum(mmsta_sep[mmsta_sep['Wire Internal Name'] == wire][PN_columns_mmsta].values, axis=0)
            if np.array_equal(pene_max, pen_mmsta):
                  n_fois_duplicate = mmsta_sep[mmsta_sep['Wire Internal Name'] == wire].shape[0]
             
                  if n_fois_duplicate:
                        row_to_duplicate = df_wire
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
                          
                              data_integrated = pd.concat([data_integrated, row_to_duplicate], ignore_index=True)
                              data_integrated = data_integrated.reset_index(drop=True)
                          
      elif df_wire.shape[0] > 1 and not matching_rows.empty:
            if matching_rows.shape[0] > df_wire.shape[0]:
                  indices_valides = get_idx(df_wire,matching_rows) #(idx_max_wire,(les indices de mmsta))
                  if len(indices_valides) > 1 :
                        for idx_val in indices_valides :
                              n_fois_duplicate = len(idx_val[1])
                              idx_row_to_duplicate = idx_val[0]
                              if n_fois_duplicate > 1:
                                    row_to_duplicate = df_wire.loc[idx_row_to_duplicate]
                                    df_wire_list = df_wire_list.drop(idx_row_to_duplicate)
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
                                    
                                          data_integrated = pd.concat([data_integrated, row_to_duplicate.to_frame().T], ignore_index=True)
                                          data_integrated = data_integrated.reset_index(drop=True)
                  
                  
            else:
                  tuple_match = find_matching_rows_for_wire(wire, df_wire_list, mmsta_sep, PN_columns_wirelist, PN_columns_mmsta)
                  for wire_idx, mmsta_idx in tuple_match:
      
                        matching_row = mmsta_sep.loc[mmsta_idx]
                        
                        # # Create a new row based on the matching row
                        new_row = df_wire_list.loc[wire_idx]
                        df_wire_list = df_wire_list.drop(wire_idx)#df_wire_list[df_wire_list['Wire Internal Name'] != wire]
                        df_wire_list = df_wire_list.reset_index(drop=True)
                        new_row[PN_columns_wirelist] = matching_row[PN_columns_mmsta].values
                        
                        level = matching_row['Level'][-1]
                        
                        if int(level) == 3:
                              new_row['SN1'] = matching_row['SN3'] 
                              new_row['SN2'] = matching_row['SN2']
                              new_row['Sup Grp'] = matching_row['SN1']
                              new_row['DES'] = matching_row['DES3']
                        elif int(level) == 2:
                              new_row['SN1'] = matching_row['SN2']
                              new_row['SN2'] = matching_row['SN1']
                              new_row['DES'] = matching_row['DES2']
                        elif int(level) == 1:
                              new_row['SN1'] = matching_row['SN1']
                              new_row['DES'] = matching_row['DES1']
                        
                        # Append the new row to df_wire_list
                        data_integrated = pd.concat([data_integrated, new_row.to_frame().T], ignore_index=True)
                        data_integrated = data_integrated.reset_index(drop=True)
 
 
data_integrated.to_excel('../Data/Updated.xlsx', index=False)                 