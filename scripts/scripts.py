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