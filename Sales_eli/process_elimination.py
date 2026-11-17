import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────
FILE_PATH   = r'C:\Users\Acer\Desktop\Dan on food\Code\Sales_eli\DDF Group - sale report FY26.xlsx'
OUTPUT_PATH = r'C:\Users\Acer\Desktop\Dan on food\Code\Sales_eli\elimination_summary_new.xlsx'
EX_RATE_DOAU = 26290   # USD → VND
EX_RATE_CAN  = 18490  # CAD → VND
NUM_FMT = '#,##0'

# ─────────────────────────────────────────────────────────────────
# ELIMINATION LISTS (confirmed by user)
# ─────────────────────────────────────────────────────────────────
MASTER_PARTIES = [
    'CÔNG TY CỔ PHẦN THỰC PHẨM AN ĐIỀN',
    'CÔNG TY CỔ PHẦN THỰC PHẨM DÂN ÔN',
    'CÔNG TY CỔ PHẦN THỰC PHẨM DÂN ÔN - CHI NHÁNH MỸ PHƯỚC 3',
    'CÔNG TY CỔ PHẦN THỰC PHẨM DÂN ÔN-CN MỸ PHƯỚC 3',
    'CÔNG TY CP THỰC PHẨM DÂN ÔN - CHI NHÁNH MỸ PHƯỚC 3',
    'DAN D PAK CO., LTD',
    'DAN ON FOODS CORPORATION-CAD',
    'DAN-D FOODS HONG KONG LIMITED',
    'DAN-D FOODS HONGKONG LTD.',
    'DAN-D FOODS LTD',
    'DAN-D FOODS LTD.',
    'DAN-D FOODS( HK LTD)-US$',
    'DAN-D-PAK USA',
    'Dan-D-Foods, Ltd-Canada',
    'DOF - Vietnam',
    'DOF My Phuoc 3'
]

ELI_LISTS = {
    'DOF': {
        'name_col':  'Name',
        'rev_col':   'Net Sales',
        'cost_col':  'Total Cost',
        'month_col': 'Month',
        'rate':      1,
        'parties': MASTER_PARTIES,
    },
    'ADF': {
        'name_col':  'Name',
        'rev_col':   'Net Sales',
        'cost_col':  'Total Cost',
        'month_col': 'Month',
        'rate':      1,
        'parties': MASTER_PARTIES,
    },
    'DOAU': {
        'name_col':  'Name',
        'rev_col':   'Amount',
        'cost_col':  'Total Cost',
        'month_col': 'Month',
        'rate':      EX_RATE_DOAU,
        'parties': MASTER_PARTIES,
    },
    'CAN': {
        'name_col':  'Customer Name',
        'rev_col':   'Net Sales',
        'cost_col':  'Total Cost',
        'month_col': 'Month',
        'rate':      EX_RATE_CAN,
        'parties': MASTER_PARTIES,
    },
}

KEYWORDS = ['DAN ON', 'DAN-D FOODS', 'DANWAY', 'DOF', 'DAN D']

# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def style_cell(cell, bold=False, fill_hex=None, num_fmt=None, align='left', italic=False):
    cell.font = Font(bold=bold, italic=italic)
    if fill_hex:
        cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type='solid')
    if num_fmt:
        cell.number_format = num_fmt
    cell.alignment = Alignment(horizontal=align)

def write_header_row(ws, row_idx, headers, fill_hex='DDEBF7', num_cols=None):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=c, value=h)
        style_cell(cell, bold=True, fill_hex=fill_hex, align='center')
    return row_idx + 1

def write_section_title(ws, row_idx, title, fill_hex='FCE4D6'):
    cell = ws.cell(row=row_idx, column=1, value=title)
    style_cell(cell, bold=True, fill_hex=fill_hex)
    return row_idx + 1

def set_num(ws, row, col, val):
    c = ws.cell(row=row, column=col, value=val)
    c.number_format = NUM_FMT
    c.alignment = Alignment(horizontal='right')

# ─────────────────────────────────────────────────────────────────
# LOAD & PROCESS DATA
# ─────────────────────────────────────────────────────────────────
def load_sheet(sheet):
    cfg = ELI_LISTS[sheet]
    df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=1)
    nc = cfg['name_col']
    df[nc] = df[nc].astype(str).str.strip()

    # Normalise month column to "YYYY-MM" string
    mc = cfg['month_col']
    df[mc] = pd.to_datetime(df[mc], errors='coerce')
    df['YearMonth'] = df[mc].dt.to_period('M').astype(str)

    rate = cfg['rate']
    df['Rev_VND']  = pd.to_numeric(df[cfg['rev_col']],  errors='coerce').fillna(0) * rate
    df['Cost_VND'] = pd.to_numeric(df[cfg['cost_col']], errors='coerce').fillna(0) * rate

    eli_mask       = df[nc].isin(cfg['parties'])
    df['IsEli']    = eli_mask

    # possible review: not eliminated but keyword match
    possible_mask  = (~eli_mask) & df[nc].str.upper().apply(
        lambda x: any(k in x for k in KEYWORDS))
    df['IsPossible'] = possible_mask
    return df


def annual_summary(df):
    total_rev  = df['Rev_VND'].sum()
    total_cost = df['Cost_VND'].sum()
    eli_rev    = df.loc[df['IsEli'], 'Rev_VND'].sum()
    eli_cost   = df.loc[df['IsEli'], 'Cost_VND'].sum()
    n_rows     = df['IsEli'].sum()
    return dict(
        n_rows=int(n_rows),
        total_rev=total_rev, total_cost=total_cost,
        eli_rev=eli_rev,     eli_cost=eli_cost,
        after_rev=total_rev - eli_rev,
        after_cost=total_cost - eli_cost,
    )


def monthly_summary(df):
    """Returns DataFrame: YearMonth | total_rev | total_cost | eli_rev | eli_cost | after_rev | after_cost"""
    grp_all = df.groupby('YearMonth')[['Rev_VND', 'Cost_VND']].sum().rename(
        columns={'Rev_VND': 'total_rev', 'Cost_VND': 'total_cost'})
    grp_eli = df[df['IsEli']].groupby('YearMonth')[['Rev_VND', 'Cost_VND']].sum().rename(
        columns={'Rev_VND': 'eli_rev', 'Cost_VND': 'eli_cost'})
    m = grp_all.join(grp_eli, how='left').fillna(0)
    m['after_rev']  = m['total_rev']  - m['eli_rev']
    m['after_cost'] = m['total_cost'] - m['eli_cost']
    m.sort_index(inplace=True)
    return m.reset_index()


def detail_by_party(df):
    nc = ELI_LISTS  # unused directly, each df already has name col
    pass  # use df directly


# ─────────────────────────────────────────────────────────────────
# WRITE SUMMARY SHEET
# ─────────────────────────────────────────────────────────────────
def write_summary_sheet(ws, all_data):
    sheets_order = ['DOF', 'ADF', 'DOAU', 'CAN']
    # Styles
    BLUE  = 'DDEBF7'
    ORANGE= 'FCE4D6'
    TOTAL_FILL = 'FFF2CC'

    row = 1
    # Title
    row = write_section_title(ws, row,
        '1. SUMMARY BY SHEET (SELLING ENTITY) - BEFORE / ELIMINATED / AFTER', ORANGE)
    hdrs = ['Sheet', 'Intercompany Rows',
            'Total Revenue (Before)', 'Total Cost (Before)',
            'Revenue Eliminated', 'Cost Eliminated',
            'Revenue (After)', 'Cost (After)']
    row = write_header_row(ws, row, hdrs, BLUE)

    grand = [0]*6
    for sh in sheets_order:
        s = all_data[sh]['annual']
        vals = [s['n_rows'], s['total_rev'], s['total_cost'],
                s['eli_rev'], s['eli_cost'], s['after_rev'], s['after_cost']]
        ws.cell(row=row, column=1, value=sh).font = Font(bold=True)
        for c, v in enumerate(vals, 2):
            set_num(ws, row, c, v)
        for i in range(6):
            grand[i] += vals[i+1]
        row += 1

    # Grand total row
    cell = ws.cell(row=row, column=1, value='GRAND TOTAL (*)')
    style_cell(cell, bold=True, fill_hex=TOTAL_FILL)
    for c, v in enumerate(grand, 2):
        cell2 = ws.cell(row=row, column=c, value=v)
        style_cell(cell2, bold=True, fill_hex=TOTAL_FILL)
        cell2.number_format = NUM_FMT
        cell2.alignment = Alignment(horizontal='right')
    row += 3

    # Per-sheet detail
    for sh in sheets_order:
        df = all_data[sh]['df']
        nc = ELI_LISTS[sh]['name_col']
        df_eli = df[df['IsEli']]

        row = write_section_title(ws, row, f'Detail - sheet "{sh}"', ORANGE)
        row = write_header_row(ws, row,
            ['Related Party', 'Revenue Eliminated', 'Cost Eliminated', 'Row Count'], BLUE)

        grp = df_eli.groupby(nc).agg(
            Rev_VND=('Rev_VND', 'sum'),
            Cost_VND=('Cost_VND', 'sum'),
            RowCount=(nc, 'count'),
        ).reset_index()

        for _, r in grp.iterrows():
            ws.cell(row=row, column=1, value=r[nc])
            set_num(ws, row, 2, r['Rev_VND'])
            set_num(ws, row, 3, r['Cost_VND'])
            ws.cell(row=row, column=4, value=int(r['RowCount']))
            row += 1
        row += 2

    # Review section
    row = write_section_title(ws, row,
        '2. REVIEW - POSSIBLE RELATED PARTIES NOT IN THE EXACT LIST (NOT ELIMINATED)', ORANGE)
    note = ws.cell(row=row, column=1,
        value="Names NOT eliminated but match keywords (DAN ON, DAN-D FOODS, DANWAY, DOF, DAN D) — please review")
    style_cell(note, italic=True)
    row += 1
    row = write_header_row(ws, row,
        ['Sheet', 'Customer Name (as in data)', 'Revenue', 'Cost', 'Row Count'], BLUE)

    any_review = False
    for sh in sheets_order:
        df = all_data[sh]['df']
        nc = ELI_LISTS[sh]['name_col']
        df_poss = df[df['IsPossible']]
        if df_poss.empty:
            continue
        any_review = True
        grp = df_poss.groupby(nc).agg(
            Rev_VND=('Rev_VND', 'sum'),
            Cost_VND=('Cost_VND', 'sum'),
            RowCount=(nc, 'count'),
        ).reset_index()
        for _, r in grp.iterrows():
            ws.cell(row=row, column=1, value=sh)
            ws.cell(row=row, column=2, value=r[nc])
            set_num(ws, row, 3, r['Rev_VND'])
            set_num(ws, row, 4, r['Cost_VND'])
            ws.cell(row=row, column=5, value=int(r['RowCount']))
            row += 1
    if not any_review:
        note2 = ws.cell(row=row, column=1,
            value='(No possible-match names found outside the exact related-party list)')
        style_cell(note2, italic=True)

    # Column widths
    ws.column_dimensions['A'].width = 55
    for col in 'BCDEFGH':
        ws.column_dimensions[col].width = 22


# ─────────────────────────────────────────────────────────────────
# WRITE MONTHLY SHEET
# ─────────────────────────────────────────────────────────────────
def write_monthly_sheet(ws, all_data):
    sheets_order = ['DOF', 'ADF', 'DOAU', 'CAN']
    BLUE   = 'DDEBF7'
    ORANGE = 'FCE4D6'
    GREEN  = 'E2EFDA'
    TOTAL_FILL = 'FFF2CC'

    hdrs = ['Entity', 'Month',
            'Total Revenue (Before)', 'Total Cost (Before)',
            'Revenue Eliminated', 'Cost Eliminated',
            'Revenue (After)', 'Cost (After)',
            'Eli Rows']

    row = 1
    row = write_section_title(ws, row,
        'MONTHLY ELIMINATION BREAKDOWN BY ENTITY', ORANGE)
    row = write_header_row(ws, row, hdrs, BLUE)

    for sh in sheets_order:
        m = all_data[sh]['monthly']
        if m.empty:
            continue

        # block title row for each entity
        cell = ws.cell(row=row, column=1, value=sh)
        style_cell(cell, bold=True, fill_hex='D9E1F2')
        for c in range(2, len(hdrs)+1):
            ws.cell(row=row, column=c).fill = PatternFill(
                start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        row += 1

        # Also need row count per month — recalculate
        df = all_data[sh]['df']
        monthly_rows = df[df['IsEli']].groupby('YearMonth').size().rename('eli_rows')

        sh_totals = [0]*6
        for _, r in m.iterrows():
            ym = r['YearMonth']
            n_eli = int(monthly_rows.get(ym, 0))
            vals  = [r['total_rev'], r['total_cost'],
                     r['eli_rev'],   r['eli_cost'],
                     r['after_rev'], r['after_cost']]
            ws.cell(row=row, column=1, value='')  # blank entity col for sub-rows
            ws.cell(row=row, column=2, value=ym)
            for c, v in enumerate(vals, 3):
                set_num(ws, row, c, v)
            ws.cell(row=row, column=9, value=n_eli)
            for i in range(6):
                sh_totals[i] += vals[i]
            row += 1

        # Subtotal for entity
        cell = ws.cell(row=row, column=1, value=f'{sh} - Subtotal')
        style_cell(cell, bold=True, fill_hex=GREEN)
        ws.cell(row=row, column=2, value='').fill = PatternFill(
            start_color=GREEN, end_color=GREEN, fill_type='solid')
        for c, v in enumerate(sh_totals, 3):
            cell2 = ws.cell(row=row, column=c, value=v)
            style_cell(cell2, bold=True, fill_hex=GREEN)
            cell2.number_format = NUM_FMT
            cell2.alignment = Alignment(horizontal='right')
        ws.cell(row=row, column=9, value=int(df['IsEli'].sum()))
        row += 2

    # Column widths
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 12
    for col in 'CDEFGHI':
        ws.column_dimensions[col].width = 22


# ─────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────
def main():
    print("Loading data...")
    all_data = {}
    for sh in ['DOF', 'ADF', 'DOAU', 'CAN']:
        df = load_sheet(sh)
        all_data[sh] = {
            'df':      df,
            'annual':  annual_summary(df),
            'monthly': monthly_summary(df),
        }
    print("Building Excel output...")

    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = 'Summary'
    write_summary_sheet(ws_summary, all_data)

    ws_monthly = wb.create_sheet('Monthly')
    write_monthly_sheet(ws_monthly, all_data)

    wb.save(OUTPUT_PATH)
    print(f"Done! File saved to: {OUTPUT_PATH}")

if __name__ == '__main__':
    main()
