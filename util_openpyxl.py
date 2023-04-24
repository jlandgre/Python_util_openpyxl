# Version 4/23/23
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side

def open_wb(sfile):
    """
    open workbook and return openpyxl wb object
    JDL 3/16/23
    """
    return openpyxl.load_workbook(sfile)

def delete_sht(wb, sht):
    """
    Delete specified sheet from a workbook
    JDL 3/16/23
    """
    if sht in wb.sheetnames: wb.remove(wb[sht])
    return wb
    
def write_df_as_wb_sht(sfile, sht, df, is_index=False):
    """
    write DataFrame to specified sheet in Excel Workbook
    (Delete previous version of sheet first to write as replacement)
    JDL 3/16/23
    """
    with pd.ExcelWriter(sfile, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name=sht, index=is_index)

def clear_worksheet(ws):
    """
    Clear Excel worksheet (ws object in an openpyxl wb object)
    JDL 4/21/23
    """
    ws.delete_rows(1, ws.max_row)
    ws.delete_cols(1, ws.max_column)
    return ws

def clear_cell(cell):
    """
    openyxl clear a cell's value, formatting and cell comment/note
    """
    cell.value = None
    cell.font = openpyxl.styles.Font()
    cell.border = openpyxl.styles.Border()
    cell.fill = openpyxl.styles.PatternFill()
    cell.number_format = 'General'
    cell.alignment = openpyxl.styles.Alignment()
    if cell.comment: cell.comment = None
    return cell

def ws_to_df(ws):
    """
    Convert an openpyxl ws to a DataFrame (range index and columns)
    """
    data = ws.values
    df = pd.DataFrame(data)
    return df

def clear_columns(ws, col1, col2):
    """
    Clear specified columns
    """
    for col in range(col1, col2+1):
        for row in range(1, ws.max_row + 1):
            clear_cell(ws.cell(row=row, column=col))
    return ws

def rng_iterator(ws, cell_home, cell_end):
    """
    Return row-wise iterator to iterate over cells in range
    specified by openpyxl home and end cells. Usage: for c in cell_iterator(xxx):
    JDL 4/21/23
    """
    row_start, col_start = cell_home.row, cell_home.column
    row_end, col_end = cell_end.row, cell_end.column

    for row in range(row_start, row_end+1):
        for col in range(col_start, col_end+1):
            cell = ws.cell(row=row, column=col)
            yield cell
            
def rng_iterator_enum(ws, cell_home, cell_end):
    """
    Return row-wise iterator with row, column enumeration to iterate 
    over cells in a range specified by openpyxl home and end cells.
    Usage: for i, j, c in cell_iterator(xxx): where i and j are the
    row and column indices of cells c returned by the generator
    JDL 4/21/23
    """
    start_row, start_col = cell_home.row, cell_home.column
    end_row, end_col,  = cell_end.row, cell_end.column, 
    for i, row in enumerate(range(start_row, end_row+1), start=1):
        for j, col in enumerate(range(start_col, end_col+1), start=1):
            cell = ws.cell(row=row, column=col)
            yield (i, j, cell)

def write_dataframe(ws, df, cell_home):
    """
    Write a DataFrame to a specific openpyxl cell on an Excel ws
    cell_home argument is ws.cell for top left data cell in Excel
    JDL 4/23/23
    """
    #Create dict of cell locations for df elements
    d_cells = set_df_openpyxl_cell_locns(ws, df, cell_home)
    
    #Write data, index and column values
    for fn in [write_df_data, write_df_index, write_df_columns]:
        ws = fn(ws, df, d_cells)
    return ws 

def set_df_openpyxl_cell_locns(ws, df, cell_home):
    """
    Set ws.cells for ranges of data, index and columns
    cell_home argument is ws.cell for top left data cell in Excel
    JDL 4/23/23
    """
    row, col = row_col(cell_home)
    d_cells = {'cell_home_data':cell_home}
    d_cells['cell_end_data'] = ws.cell(row + df.index.size - 1, col + df.columns.size - 1)
    d_cells['cell_home_idx'] = ws.cell(row, col - 1)
    d_cells['cell_end_idx'] = ws.cell(row + df.index.size - 1, col - 1)
    d_cells['cell_home_cols'] = ws.cell(row - 1, col)
    d_cells['cell_end_cols'] = ws.cell(row - 1, col + df.columns.size - 1)    
    return d_cells

def row_col(c):
    """
    return openpyxl ws.cell row and column tuple
    JDL 4/23/23
    """
    return c.row, c.column 
    
def write_df_data(ws, df, d_cells):
    """
    Write DataFrame's data values
    JDL 4/23/23
    """
    for i, j, c in rng_iterator_enum(ws, d_cells['cell_home_data'], d_cells['cell_end_data']):
        c.value = df.values[i-1][j-1]        
    return ws
    
def write_df_index(ws, df, d_cells):
    """
    Write DataFrame's index to column adjacent to first data column
    JDL 4/23/23
    """
    #Write index values
    for i, j, c in rng_iterator_enum(ws, d_cells['cell_home_idx'], d_cells['cell_end_idx']):
        c.value = list(df.index)[i-1]
    
    #Write index name as heading above index values
    ws.cell(d_cells['cell_home_idx'].row - 1, d_cells['cell_home_idx'].column).value = df.index.name
    return ws

def write_df_columns(ws, df, d_cells):
    """
    Write DataFrame's column values to row above to first data row
    JDL 4/23/23
    """    
    for i, j, c in rng_iterator_enum(ws, d_cells['cell_home_cols'], d_cells['cell_end_cols']):
        c.value = list(df.columns)[j-1]  
    return ws

def set_openpyxl_border_obj(style_border):
    """
    Create a border style based on style_border='thick', 'thin' etc.
    Use "from openpyxl.styles import Border, Side" to import needed openpyxl attributes
    JDL 4/21/23
    """
    return Border(left=Side(style=style_border),
                  right=Side(style=style_border), 
                  top=Side(style=style_border),
                  bottom=Side(style=style_border))

def set_range_border(ws, cell_home, cell_end, style_border):
    #Create a Border object for style_border
    border_obj = set_openpyxl_border_obj(style_border)
    
    #Apply the border_obj to each cell in the range
    for c in rng_iterator(ws, cell_home, cell_end):
        c.border = border_obj
        
def set_df_borders(ws, df, cell_home):
    d_cells = set_df_openpyxl_cell_locns(ws, df, cell_home)
    ws = set_df_data_borders(ws, d_cells, 'thin')
    ws = set_df_index_borders(ws, d_cells, 'thick')
    ws = set_df_cols_borders(ws, d_cells, 'thick')
    return ws

def set_df_data_borders(ws, d_cells, style_border):
    """
    Put border around cells for df data values
    """
    set_range_border(ws, d_cells['cell_home_data'], d_cells['cell_end_data'], style_border)
    return ws

def set_df_index_borders(ws, d_cells, style_border):
    set_range_border(ws, d_cells['cell_home_idx'], d_cells['cell_end_idx'], style_border)
    row = d_cells['cell_home_idx'].row - 1
    col = d_cells['cell_home_idx'].column
    set_range_border(ws, ws.cell(row, col), ws.cell(row, col), style_border)
    return ws

def set_df_cols_borders(ws, d_cells, style_border):
    set_range_border(ws, d_cells['cell_home_cols'], d_cells['cell_end_cols'], style_border)
    return ws

