"""Build the phoning prospect workbook from parsed.json."""
import json, re, sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as CL

SRC, OUT = sys.argv[1], sys.argv[2]
data = json.load(open(SRC))
# A company name ending with ',' continues on the next (regular-weight) line in the PDF.
for _p in data:
    for _c in _p['contacts']:
        if _c['societe'].endswith(',') and _c['adresse']:
            _c['societe'] = _c['societe'] + ' ' + _c['adresse'].pop(0)

# ------------------------------------------------------------------ translations (German source → English)
STADE = {'Baubewilligung erteilt': 'Permit granted', 'Baugesuch eingereicht': 'Permit application filed'}
EVENT = {'Baubewilligung erteilt': 'Permit granted', 'Baugesuch publiziert': 'Application published'}
ROLE = {'Bauherr': 'Owner / client', 'Architekt/Planer': 'Architect / planner',
        'Generalunternehmer': 'General contractor', 'Bauingenieur': 'Civil engineer',
        'Bauherr / Architekt/Planer': 'Owner & architect'}
ROLE_ORDER = {'Owner / client': 0, 'Owner & architect': 0, 'Architect / planner': 1,
              'General contractor': 2, 'Civil engineer': 3}
SOURCE = {'fiche': 'Detail card', 'liste complémentaire': 'Summary table'}
KAT = {'Handel und Verwaltung': 'Commerce and administration',
       'Unterricht, Bildung und Forschung': 'Education and research',
       'Industrie und Gewerbe': 'Industry and trade', 'Land- und Forstwirtschaft': 'Agriculture and forestry',
       'Kultur und Geselligkeit': 'Culture and leisure', 'Fürsorge und Gesundheit': 'Social care and health',
       'Militär- und Schutzanlagen': 'Military and civil protection'}
ZWECK = {'Nicht relevant': 'Not relevant', 'Eigenbedarf': 'Own use'}
WORK = {'Umbau innen': 'interior conversion', 'Umbau aussen': 'exterior conversion',
        'Neubau': 'new build', 'Anbau': 'extension', 'Abbruch': 'demolition'}
TYPE = {
    'Restaurationsbetriebe': 'Restaurants', 'Wohnungen': 'Housing',
    'Aussenanlagen, Kinderspielplätze und Parkanlagen': 'Outdoor areas, playgrounds and parks',
    'Fitnesscenter/-raum': 'Fitness centre / gym', 'Parkhäuser und Einstellhallen': 'Car parks and underground garages',
    'Garagen / Fertiggaragen': 'Garages / prefabricated garages', 'Feuerwehrgebäude': 'Fire stations',
    'Konzertbauten und Theaterbauten': 'Concert halls and theatres', 'Strassen': 'Roads',
    'Primar- und Sekundarschulen': 'Primary and secondary schools',
    'Berufs- und höhere Fachschulen': 'Vocational schools and universities of applied sciences',
    'Mittelschulen und Gymnasien': 'Upper secondary schools', 'Hochschulen und Universitäten': 'Universities',
    'Forschungsinstitute': 'Research institutes', 'Lagerhallen': 'Warehouses', 'Industriehallen': 'Industrial halls',
    'Industrielle Produktionsbauten': 'Industrial production buildings',
    'Betriebs- und Gewerbebauten': 'Operational and commercial buildings', 'Atelier und Studio': 'Workshops and studios',
    'Schuppen und Hütten': 'Sheds and huts', 'Futterlagerräume, Treibhäuser und Silobauten': 'Feed stores, greenhouses and silos',
    'Stallungen und landwirtschaftliche Produktionsanlagen': 'Stables and agricultural facilities',
    'Tierspitäler': 'Veterinary clinics', 'Jauchegrube': 'Slurry pits', 'Ladenbauten': 'Shops',
    'Bürobauten mit einfachen Anforderungen': 'Offices (standard specification)',
    'Bürobauten mit erhöhten Anforderungen': 'Offices (high specification)',
    'Verwaltungsgebäude und Rechenzentren': 'Administrative buildings and data centres',
    'Banken, Postgebäude und Fernmeldegebäude': 'Banks, post and telecom buildings',
    'Gemeindehäuser, Rathäuser und Regierungsgebäude': 'Town halls and government buildings',
    'Wiedereingliederungsstätten': 'Rehabilitation centres', 'Arztpraxen und Ärztehäuser': 'Medical practices',
}


def tr_sub(s):
    m = re.match(r'^(\d+)\s+(.+?)\s+\((.+)\)$', s)
    if not m:
        return s
    code, typ, work = m.groups()
    return f"{TYPE.get(typ, typ)} – {WORK.get(work, work)}"


def civ(name):
    return re.sub(r'^Herr\b', 'Mr', re.sub(r'^Frau\b', 'Ms', name))


def to_date(s):
    return datetime.strptime(s, '%d.%m.%Y') if s else None


def montant(s):
    m = re.match(r'([\d.]+)\s*Mio', s or '')
    return round(float(m.group(1)) * 1_000_000) if m else None


def tel_link(t):
    digits = re.sub(r'\D', '', t)
    return 'tel:+41' + digits[1:] if digits.startswith('0') else None


def nat_key(name):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r'(\d+)', name)]


# ------------------------------------------------------------------ duplicates (same Infopro object in 2 files)
by_obj = {}
for p in data:
    by_obj.setdefault(p['kv']['Objektnummer'], []).append(f"{p['fichier']} (p.{p['page']})")


def dup_of(p):
    others = [x for x in by_obj[p['kv']['Objektnummer']] if x != f"{p['fichier']} (p.{p['page']})"]
    return 'Same project as ' + ', '.join(others) if others else ''


# ------------------------------------------------------------------ styles
F = 'Arial'
font = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
hfont = Font(name=F, size=10, bold=True, color='FFFFFF')
title_font = Font(name=F, size=16, bold=True, color='0E3A40')
h2 = Font(name=F, size=12, bold=True, color='0E3A40')
muted = Font(name=F, size=9, italic=True, color='5C6B6D')
link = Font(name=F, size=10, color='1C5F68', underline='single')
missing = Font(name=F, size=10, italic=True, color='B4541A')
FILL_H = PatternFill('solid', fgColor='0E3A40')      # data header
FILL_H_IN = PatternFill('solid', fgColor='B4541A')   # header of columns to fill in
FILL_H_DET = PatternFill('solid', fgColor='5C6B6D')  # header of detail columns
FILL_IN = PatternFill('solid', fgColor='FFF2CC')     # cells to fill in (yellow)
FILL_SOFT = PatternFill('solid', fgColor='EAF2F2')
thin = Side(style='thin', color='DEE5E3')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')
TOP = Alignment(vertical='top')
CENTER = Alignment(horizontal='center', vertical='top')
DATE_FMT = 'dd/mm/yyyy'
CHF_FMT = "#,##0 \"CHF\""

wb = Workbook()

# ------------------------------------------------------------------ Lists (dropdowns + parameters)
ls = wb.active
ls.title = 'Lists'
LISTS = {
    'Call result': ['Reached – right person', 'Reached – gatekeeper / switchboard', 'Voicemail (message left)',
                    'Voicemail (no message)', 'No answer', 'Busy', 'Wrong / invalid number'],
    'Status': ['To call', 'Call back', 'Interested', 'Meeting booked', 'Quote requested', 'Not interested',
               'Already has a supplier', 'Not a fit', 'Project completed / cancelled'],
    'Next action': ['Call back', 'Send brochure', 'Send quote', 'Meeting / site visit',
                    'Follow-up email', 'None (closed)'],
}
for j, (name, vals) in enumerate(LISTS.items()):
    c = 1 + j * 2
    ls.cell(1, c, name).font = bold
    for i, v in enumerate(vals, 2):
        ls.cell(i, c, v).font = font
    ls.column_dimensions[CL(c)].width = 34
    ls.column_dimensions[CL(c + 1)].width = 3
# priority thresholds
ls['H1'] = 'Priority settings'; ls['H1'].font = bold
ls['H2'] = 'Priority A threshold (CHF ≥)'; ls['I2'] = 300000
ls['H3'] = 'Priority B threshold (CHF ≥)'; ls['I3'] = 100000
ls['H4'] = 'Below the B threshold → priority C'
for r in (2, 3, 4):
    ls[f'H{r}'].font = font
for r in (2, 3):
    ls[f'I{r}'].font = Font(name=F, size=10, color='0000FF')
    ls[f'I{r}'].fill = FILL_IN
    ls[f'I{r}'].number_format = CHF_FMT
ls['H6'] = ('Suggested default thresholds (adjust as needed), based on the works value stated on the Infopro card. '
            'Editing the yellow cells recalculates the "Priority" column on the Calls sheet.')
ls['H6'].font = muted
ls['H6'].alignment = Alignment(wrap_text=True, vertical='top')
ls.merge_cells('H6:I9')
ls.column_dimensions['H'].width = 34
ls.column_dimensions['I'].width = 16


def list_ref(name):
    j = list(LISTS).index(name)
    col = CL(1 + j * 2)
    return f"Lists!${col}$2:${col}${len(LISTS[name]) + 1}"


# ------------------------------------------------------------------ Calls (one row per contact × project)
ws = wb.create_sheet('Calls', 0)
COLS = [
    # (header, width, group)  group: core | input | detail
    ('PDF file', 21, 'core'), ('Page', 6, 'core'), ('Priority', 8, 'core'), ('Role', 17, 'core'),
    ('Company', 34, 'core'), ('Contact person', 20, 'core'), ('Phone', 15, 'core'), ('Email', 26, 'core'),
    ('Project', 34, 'core'), ('Site location', 18, 'core'), ('Stage', 16, 'core'),
    ('Works value (CHF)', 14, 'core'), ('Linked projects (same contact)', 10, 'core'),
    ('Last call result', 22, 'input'), ('Status', 20, 'input'), ('Attempts', 9, 'input'),
    ('Last call date', 12, 'input'), ('Next action', 19, 'input'), ('Follow-up date', 12, 'input'),
    ('Notes / person reached', 40, 'input'),
    ('Infopro project ref.', 11, 'detail'), ('Stage date', 11, 'detail'), ('Contact address', 30, 'detail'),
    ('Website', 22, 'detail'), ('Site address', 28, 'detail'), ('Category', 22, 'detail'),
    ('Project description', 50, 'detail'), ('Contact source', 14, 'detail'), ('Duplicate', 24, 'detail'),
]
H = {name: i + 1 for i, (name, _, _) in enumerate(COLS)}
for (name, width, grp), i in zip(COLS, range(1, len(COLS) + 1)):
    c = ws.cell(1, i, name)
    c.font = hfont
    c.fill = {'core': FILL_H, 'input': FILL_H_IN, 'detail': FILL_H_DET}[grp]
    c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    c.border = BORDER
    ws.column_dimensions[CL(i)].width = width
ws.row_dimensions[1].height = 42

rows = []
for p in data:
    last = p['termine'][-1] if p['termine'] else {'date': '', 'evenement': ''}
    for k, c in enumerate(p['contacts']):
        rows.append((p, c, k))
rows.sort(key=lambda t: (nat_key(t[0]['fichier']), t[0]['page'], ROLE_ORDER.get(ROLE.get(t[1]['role'], ''), 9), t[2]))

r = 1
for p, c, _ in rows:
    r += 1
    last = p['termine'][-1] if p['termine'] else {'date': ''}
    role = ROLE.get(c['role'], c['role'])
    tel = c['tel']
    vals = {
        'PDF file': p['fichier'], 'Page': p['page'], 'Role': role, 'Company': c['societe'],
        'Contact person': ', '.join(civ(x) for x in c['interlocuteurs']),
        'Phone': tel or 'To be found', 'Email': c['email'],
        'Project': p['projet'], 'Site location': p['lieu_titre'], 'Stage': STADE.get(p['stade'], p['stade']),
        'Works value (CHF)': montant(p['kv'].get('Bausumme')),
        'Status': 'To call', 'Attempts': 0,
        'Infopro project ref.': p['kv'].get('Objektnummer'), 'Stage date': to_date(last['date']),
        'Contact address': ', '.join(c['adresse']), 'Website': c['web'],
        'Site address': p['adresse_chantier'], 'Category': KAT.get(p['kv'].get('Kategorie'), p['kv'].get('Kategorie')),
        'Project description': p['description'], 'Contact source': SOURCE.get(c['source'], c['source']), 'Duplicate': dup_of(p),
    }
    for name, v in vals.items():
        ws.cell(r, H[name], v)
    # formulas
    cm, cp, ct, cs = CL(H['Works value (CHF)']), CL(H['Priority']), CL(H['Phone']), CL(H['Company'])
    ws.cell(r, H['Priority'], f'=IF({cm}{r}="","C",IF({cm}{r}>=Lists!$I$2,"A",IF({cm}{r}>=Lists!$I$3,"B","C")))')
    ws.cell(r, H['Linked projects (same contact)'],
            f'=IF({ct}{r}="To be found",COUNTIF(${cs}$2:${cs}$9999,{cs}{r}),COUNTIF(${ct}$2:${ct}$9999,{ct}{r}))')
    # hyperlinks
    if tel and tel_link(tel):
        ws.cell(r, H['Phone']).hyperlink = tel_link(tel)
    if c['email']:
        ws.cell(r, H['Email']).hyperlink = 'mailto:' + c['email']
    if c['web']:
        ws.cell(r, H['Website']).hyperlink = c['web']
LAST = r

# cell styling
for row in ws.iter_rows(min_row=2, max_row=LAST):
    for cell in row:
        name = COLS[cell.column - 1][0]
        grp = COLS[cell.column - 1][2]
        cell.font = font
        cell.border = BORDER
        cell.alignment = WRAP if name in ('Project', 'Company', 'Notes / person reached', 'Role') else TOP
        if grp == 'input':
            cell.fill = FILL_IN
        if name in ('Page', 'Priority', 'Attempts', 'Linked projects (same contact)'):
            cell.alignment = CENTER
        if name in ('Last call date', 'Follow-up date', 'Stage date'):
            cell.number_format = DATE_FMT
        if name == 'Works value (CHF)':
            cell.number_format = CHF_FMT
        if name == 'Priority':
            cell.font = bold
        if name in ('Phone', 'Email', 'Website') and cell.hyperlink:
            cell.font = link
        if name == 'Phone' and cell.value == 'To be found':
            cell.font = missing
        if name == 'Company':
            cell.font = bold

# table, freeze, grouping
ref = f"A1:{CL(len(COLS))}{LAST}"
tbl = Table(displayName='tblCalls', ref=ref)
tbl.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=False)
ws.add_table(tbl)
ws.freeze_panes = 'F2'
first_det, last_det = H['Infopro project ref.'], len(COLS)
for ci in range(first_det, last_det + 1):  # per-column outline keeps each column's own width
    ws.column_dimensions[CL(ci)].outlineLevel = 1
for rr in range(2, LAST + 1):
    ws.row_dimensions[rr].height = 30  # fixed 2-line rows: the call sheet stays scannable

# data validation
def add_dv(col_name, formula):
    dv = DataValidation(type='list', formula1=formula, allow_blank=True, showErrorMessage=True,
                        errorTitle='Unexpected value', error='Pick a value from the list (Lists sheet).')
    ws.add_data_validation(dv)
    col = CL(H[col_name])
    dv.add(f"{col}2:{col}{LAST + 500}")

add_dv('Last call result', list_ref('Call result'))
add_dv('Status', list_ref('Status'))
add_dv('Next action', list_ref('Next action'))
for nm in ('Last call date', 'Follow-up date'):
    dv = DataValidation(type='date', operator='greaterThan', formula1='DATE(2020,1,1)', allow_blank=True,
                        showErrorMessage=True, errorTitle='Invalid date', error='Enter a date (dd/mm/yyyy).')
    ws.add_data_validation(dv)
    dv.add(f"{CL(H[nm])}2:{CL(H[nm])}{LAST + 500}")
dv = DataValidation(type='whole', operator='between', formula1='0', formula2='50', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"{CL(H['Attempts'])}2:{CL(H['Attempts'])}{LAST + 500}")

# conditional formatting
q, rl, pr = CL(H['Status']), CL(H['Follow-up date']), CL(H['Priority'])
rng_all = f"A2:{CL(len(COLS))}{LAST + 500}"
# whole row tint by qualification (soft), keyed on the qualification column
ws.conditional_formatting.add(rng_all, FormulaRule(formula=[f'OR(${q}2="Meeting booked",${q}2="Quote requested")'],
                              fill=PatternFill('solid', fgColor='C6EFCE')))
ws.conditional_formatting.add(rng_all, FormulaRule(formula=[f'${q}2="Interested"'],
                              fill=PatternFill('solid', fgColor='E2F0D9')))
ws.conditional_formatting.add(rng_all, FormulaRule(
    formula=[f'OR(${q}2="Not interested",${q}2="Not a fit",${q}2="Project completed / cancelled",${q}2="Already has a supplier")'],
    font=Font(color='8C8C8C'), fill=PatternFill('solid', fgColor='F2F2F2')))
# overdue follow-up (date passed and file still open)
ws.conditional_formatting.add(f"{rl}2:{rl}{LAST + 500}", FormulaRule(
    formula=[f'AND({rl}2<>"",{rl}2<=TODAY(),OR(${q}2="Call back",${q}2="Interested",${q}2="Quote requested",${q}2="To call"))'],
    font=Font(name=F, bold=True, color='9C0006'), fill=PatternFill('solid', fgColor='FFC7CE')))
ws.conditional_formatting.add(f"{pr}2:{pr}{LAST + 500}", CellIsRule(operator='equal', formula=['"A"'],
                              font=Font(name=F, bold=True, color='FFFFFF'), fill=PatternFill('solid', fgColor='B4541A')))
ws.conditional_formatting.add(f"{pr}2:{pr}{LAST + 500}", CellIsRule(operator='equal', formula=['"B"'],
                              fill=PatternFill('solid', fgColor='FCE4D6')))

# header comments (guidance)
notes = {
    'Priority': 'Based on the works value (thresholds can be changed on the Lists sheet).',
    'Linked projects (same contact)': 'Number of rows with the same phone number (or the same company when there is no number). '
                                      'If > 1, one call can cover several projects.',
    'Last call result': 'What happened on the call itself (reached, voicemail, etc.).',
    'Status': 'Where the prospect stands commercially. Defaults to "To call".',
    'Follow-up date': 'Always set a follow-up date: it turns red once it is overdue.',
    'Phone': 'Clickable: starts the call in your softphone / Teams / Skype (tel:+41... link).',
    'Contact source': '"Detail card" = full contact block; "Summary table" = short table at the bottom of the card.',
}
for k, v in notes.items():
    ws.cell(1, H[k]).comment = Comment(v, 'Prospecting')

# ------------------------------------------------------------------ Projects (one row per project)
wp = wb.create_sheet('Projects', 1)
PCOLS = [('PDF file', 21), ('Page', 6), ('Infopro project ref.', 11), ('Project', 38), ('Site location', 18),
         ('Site address', 28), ('Plot', 12), ('Stage', 18), ('Application published', 12),
         ('Permit granted', 12), ('Works value (CHF)', 14), ('Category', 24), ('Use', 13),
         ('Subcategories', 42), ('Description', 55), ('Infopro research date', 12),
         ('Owner / client', 28), ('Architect / planner', 28), ('Contacts', 9),
         ('Contacts reached', 9), ('Interested / meetings / quotes', 10), ('Duplicate', 26)]
PH = {n: i + 1 for i, (n, _) in enumerate(PCOLS)}
for i, (n, w) in enumerate(PCOLS, 1):
    c = wp.cell(1, i, n)
    c.font, c.fill, c.border = hfont, FILL_H, BORDER
    c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    wp.column_dimensions[CL(i)].width = w
wp.row_dimensions[1].height = 42
A = {n: CL(i) for n, i in H.items()}  # column letters in Calls
projects = sorted(data, key=lambda p: (nat_key(p['fichier']), p['page']))
for r, p in enumerate(projects, 2):
    ev = {EVENT.get(t['evenement'], t['evenement']): to_date(t['date']) for t in p['termine']}
    mo = [c['societe'] for c in p['contacts'] if c['role'].startswith('Bauherr') and c['source'] == 'fiche']
    ar = [c['societe'] for c in p['contacts'] if 'Architekt' in c['role'] and c['source'] == 'fiche']
    vals = {'PDF file': p['fichier'], 'Page': p['page'], 'Infopro project ref.': p['kv'].get('Objektnummer'),
            'Project': p['projet'], 'Site location': p['lieu_titre'], 'Site address': p['adresse_chantier'],
            'Plot': p['reference'], 'Stage': STADE.get(p['stade'], p['stade']),
            'Application published': ev.get('Application published'), 'Permit granted': ev.get('Permit granted'),
            'Works value (CHF)': montant(p['kv'].get('Bausumme')),
            'Category': KAT.get(p['kv'].get('Kategorie'), p['kv'].get('Kategorie')),
            'Use': ZWECK.get(p['kv'].get('Verwendungszweck'), p['kv'].get('Verwendungszweck') or ''),
            'Subcategories': '\n'.join(tr_sub(s) for s in p['sous_categories']),
            'Description': p['description'], 'Infopro research date': to_date(p['kv'].get('Recherchedatum')),
            'Owner / client': ' / '.join(mo), 'Architect / planner': ' / '.join(ar), 'Duplicate': dup_of(p)}
    for n, v in vals.items():
        wp.cell(r, PH[n], v)
    key = f'$A{r}'
    f_pdf, f_pg = f"Calls!${A['PDF file']}$2:${A['PDF file']}$9999", f"Calls!${A['Page']}$2:${A['Page']}$9999"
    f_res, f_q = (f"Calls!${A['Last call result']}$2:${A['Last call result']}$9999",
                  f"Calls!${A['Status']}$2:${A['Status']}$9999")
    wp.cell(r, PH['Contacts'], f'=COUNTIFS({f_pdf},{key},{f_pg},$B{r})')
    wp.cell(r, PH['Contacts reached'],
            f'=COUNTIFS({f_pdf},{key},{f_pg},$B{r},{f_res},"Reached*")')
    wp.cell(r, PH['Interested / meetings / quotes'],
            f'=COUNTIFS({f_pdf},{key},{f_pg},$B{r},{f_q},"Interested")'
            f'+COUNTIFS({f_pdf},{key},{f_pg},$B{r},{f_q},"Meeting booked")'
            f'+COUNTIFS({f_pdf},{key},{f_pg},$B{r},{f_q},"Quote requested")')
PLAST = len(projects) + 1
for row in wp.iter_rows(min_row=2, max_row=PLAST):
    for cell in row:
        n = PCOLS[cell.column - 1][0]
        cell.font, cell.border = font, BORDER
        cell.alignment = WRAP if n in ('Project', 'Description', 'Subcategories', 'Site address', 'Owner / client',
                                       'Architect / planner', 'Duplicate', 'Category') else TOP
        if n in ('Application published', 'Permit granted', 'Infopro research date'):
            cell.number_format = DATE_FMT
        if n == 'Works value (CHF)':
            cell.number_format = CHF_FMT
        if n in ('Page', 'Contacts', 'Contacts reached', 'Interested / meetings / quotes'):
            cell.alignment = CENTER
        if n == 'Project':
            cell.font = bold
t2 = Table(displayName='tblProjects', ref=f"A1:{CL(len(PCOLS))}{PLAST}")
t2.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=True)
wp.add_table(t2)
wp.freeze_panes = 'E2'
wp.conditional_formatting.add(f"V2:V{PLAST}", FormulaRule(formula=['V2<>""'], fill=PatternFill('solid', fgColor='FCE4D6')))
wp.conditional_formatting.add(f"{CL(PH['Interested / meetings / quotes'])}2:{CL(PH['Interested / meetings / quotes'])}{PLAST}",
                              CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='C6EFCE'),
                                         font=Font(name=F, bold=True)))

# ------------------------------------------------------------------ Dashboard
wd = wb.create_sheet('Dashboard', 0)
wd.sheet_view.showGridLines = False
wd.column_dimensions['A'].width = 3
wd.column_dimensions['B'].width = 38
wd.column_dimensions['C'].width = 14
wd.column_dimensions['D'].width = 12
wd.column_dimensions['E'].width = 4
wd.column_dimensions['F'].width = 36
wd.column_dimensions['G'].width = 12
wd.column_dimensions['H'].width = 12
wd['B2'] = 'Phone prospecting — construction projects (Infopro "DRINGEND" cards)'
wd['B2'].font = title_font
wd['B3'] = (f"{len(set(p['fichier'] for p in data))} PDF files · {len(data)} projects · {LAST - 1} contacts. "
            "The figures below update automatically from the Calls sheet.")
wd['B3'].font = muted

ap = lambda n: f"Calls!${A[n]}$2:${A[n]}$9999"
def kpi(row, label, formula, fmt=None, note=None):
    wd.cell(row, 2, label).font = font
    c = wd.cell(row, 3, formula)
    c.font = Font(name=F, size=11, bold=True, color='0E3A40')
    c.alignment = Alignment(horizontal='right')
    if fmt:
        c.number_format = fmt
    if note:
        wd.cell(row, 4, note).font = muted
    for col in (2, 3):
        wd.cell(row, col).border = Border(bottom=thin)

wd['B5'] = 'Overview'; wd['B5'].font = h2
kpi(6, 'Projects', f"=COUNTA(Projects!$A$2:$A${PLAST})")
kpi(7, 'Contacts to call', f"=COUNTA({ap('Company')})")
kpi(8, 'Contacts without a phone number', f'=COUNTIF({ap("Phone")},"To be found")')
kpi(9, 'Total works value', f"=SUM(Projects!$K$2:$K${PLAST})", CHF_FMT, 'duplicates included')
wd['B11'] = 'Call progress'; wd['B11'].font = h2
kpi(12, 'Contacts called (≥ 1 attempt)', f'=COUNTIF({ap("Attempts")},">0")')
kpi(13, 'Coverage rate', '=IFERROR(C12/C7,0)', '0%')
kpi(14, 'Total attempts', f'=SUM({ap("Attempts")})')
kpi(15, 'Contacts reached', f'=COUNTIF({ap("Last call result")},"Reached*")')
kpi(16, 'Reach rate', '=IFERROR(C15/C12,0)', '0%', 'reached / called')
kpi(17, 'Interested + meetings + quotes', f'=COUNTIF({ap("Status")},"Interested")+COUNTIF({ap("Status")},"Meeting booked")+COUNTIF({ap("Status")},"Quote requested")')
kpi(18, 'Conversion rate', '=IFERROR(C17/C15,0)', '0%', 'qualified / reached')
kpi(19, 'Overdue follow-ups',
    f'=SUMPRODUCT(({ap("Follow-up date")}<>"")*({ap("Follow-up date")}<=TODAY())*'
    f'(({ap("Status")}="Call back")+({ap("Status")}="Interested")+({ap("Status")}="Quote requested")+({ap("Status")}="To call")))')
wd['C19'].font = Font(name=F, size=11, bold=True, color='9C0006')

wd['F5'] = 'Contacts by status'; wd['F5'].font = h2
wd['F6'], wd['G6'] = 'Status', 'Contacts'
for c in ('F6', 'G6'):
    wd[c].font, wd[c].fill = hfont, FILL_H
for i, v in enumerate(LISTS['Status'], 7):
    wd.cell(i, 6, v).font = font
    wd.cell(i, 7, f'=COUNTIF({ap("Status")},F{i})').font = font
    for col in (6, 7):
        wd.cell(i, col).border = Border(bottom=thin)
end_q = 6 + len(LISTS['Status'])
wd.cell(end_q + 1, 6, 'Total').font = bold
wd.cell(end_q + 1, 7, f'=SUM(G7:G{end_q})').font = bold

row0 = end_q + 3
wd.cell(row0, 6, 'Contacts by priority').font = h2
wd.cell(row0 + 1, 6, 'Priority').font = hfont; wd.cell(row0 + 1, 6).fill = FILL_H
wd.cell(row0 + 1, 7, 'Contacts').font = hfont; wd.cell(row0 + 1, 7).fill = FILL_H
wd.cell(row0 + 1, 8, 'Still to call').font = hfont; wd.cell(row0 + 1, 8).fill = FILL_H
for i, v in enumerate(['A', 'B', 'C'], row0 + 2):
    wd.cell(i, 6, v).font = bold
    wd.cell(i, 7, f'=COUNTIF({ap("Priority")},F{i})').font = font
    wd.cell(i, 8, f'=COUNTIFS({ap("Priority")},F{i},{ap("Status")},"To call")').font = font

row1 = row0 + 6
wd.cell(row1, 6, 'Projects by stage').font = h2
for col, t in ((6, 'Stage'), (7, 'Projects')):
    wd.cell(row1 + 1, col, t).font = hfont
    wd.cell(row1 + 1, col).fill = FILL_H
for i, v in enumerate(sorted(set(STADE.values())), row1 + 2):
    wd.cell(i, 6, v).font = font
    wd.cell(i, 7, f'=COUNTIF(Projects!$H$2:$H${PLAST},F{i})').font = font

tips_row = row1 + 6
wd.cell(tips_row, 2, 'Reminders').font = h2
tips = [
    '1. On the Calls sheet, filter Status = "To call" and sort by Priority.',
    '2. Fill in the yellow columns straight after each call (result, status, date, follow-up).',
    '3. Never leave a row open without a follow-up date: overdue follow-ups turn red.',
    '4. "Linked projects" > 1: the same person handles several projects, so cover them in one call.',
    '5. Open the PDF named in column A for the full project card.',
]
for i, t in enumerate(tips, tips_row + 1):
    wd.cell(i, 2, t).font = font

# ------------------------------------------------------------------ How to use
wm = wb.create_sheet('How to use', 1)
wm.sheet_view.showGridLines = False
wm.column_dimensions['A'].width = 3
wm.column_dimensions['B'].width = 30
wm.column_dimensions['C'].width = 95
r = 2
wm.cell(r, 2, 'How to use this prospecting file').font = title_font
r += 2
sections = [
    ('Sheets', [
        ('Dashboard', 'Automatically calculated indicators (progress, reach rate, conversion, overdue follow-ups).'),
        ('Calls', 'The call list: one row per contact and per project. This is where you work.'),
        ('Projects', 'One row per project (= one PDF page), with all the data extracted from the card.'),
        ('Lists', 'Drop-down values and priority thresholds (editable).'),
    ]),
    ('Colour code', [
        ('Dark teal header', 'Data extracted from the PDFs (do not edit, except to correct).'),
        ('Orange header + yellow cells', 'Columns to fill in while calling.'),
        ('Grey header', 'Extra details: grouped columns, click the "−" above them to hide them.'),
        ('Green row', 'Prospect interested, meeting booked or quote requested.'),
        ('Greyed-out row', 'Closed (not interested, not a fit, already has a supplier, project completed).'),
        ('Red follow-up date', 'Overdue follow-up on a row that is still open: deal with it first.'),
        ('"To be found" (orange)', 'No number on the card: look it up on local.ch / search.ch / the company website.'),
    ]),
    ('Columns to fill in', [
        ('Last call result', 'What happened: reached (right person or gatekeeper), voicemail, no answer, busy, wrong number.'),
        ('Status', 'Where the prospect stands: to call → call back → interested → meeting / quote; or closed.'),
        ('Attempts', 'Add one per call (good practice: 5 to 6 attempts at most before closing).'),
        ('Last call date', 'Date of the last call (Ctrl + ; inserts today\'s date).'),
        ('Next action / Follow-up date', 'Always an action and a date, unless the row is closed.'),
        ('Notes', 'Name and role of the person reached, stated need, objection, best time to call back.'),
    ]),
    ('Example of a filled-in row', [
        ('Last call result', 'Reached – gatekeeper / switchboard'),
        ('Status', 'Call back'),
        ('Attempts', '2'),
        ('Last call date', '14/10/2026'),
        ('Next action', 'Call back'),
        ('Follow-up date', '16/10/2026'),
        ('Notes', 'Switchboard: Mr Franco (facilities manager) on a site visit, call back Thursday after 2pm. 3rd-floor project confirmed.'),
    ]),
    ('Glossary (cards are in German)', [
        ('Owner / client (Bauherr)', 'The end client: the owner or tenant commissioning the works.'),
        ('Architect / planner', 'The firm that designs the project and often specifies the suppliers.'),
        ('Permit application filed', 'Baugesuch eingereicht: project awaiting approval, works still to come (earliest point to get in).'),
        ('Permit granted', 'Baubewilligung erteilt: approval obtained, works starting soon (urgent).'),
        ('Infopro project ref.', 'Objektnummer: unique project ID at Infopro Digital (Baublatt).'),
    ]),
    ('Notes on the data', [
        ('Source', 'Automatic extraction of the 91 PDFs in the "harworth" folder (Infopro Digital Schweiz cards).'),
        ('Original language', 'Company names, addresses, project titles and descriptions are kept as they appear in the PDFs (mostly French).'),
        ('Multi-page files', 'Some PDFs contain several projects: the "Page" column gives the project\'s page.'),
        ('Duplicates', '3 projects appear in two different files ("Duplicate" column): call them only once.'),
        ('Summary table', '5 contacts come from the short table at the bottom of the card (name, town and phone only).'),
        ('Works values', 'Works value as stated by Infopro (an estimate, converted from "Mio CHF").'),
    ]),
]
for title, items in sections:
    wm.cell(r, 2, title).font = h2
    r += 1
    for k, v in items:
        a = wm.cell(r, 2, k)
        b = wm.cell(r, 3, v)
        a.font, b.font = bold, font
        a.alignment = b.alignment = Alignment(wrap_text=True, vertical='top')
        if title == 'Example of a filled-in row':
            b.fill = FILL_IN
        for c in (a, b):
            c.border = Border(bottom=thin)
        r += 1
    r += 1

# ------------------------------------------------------------------ workbook-wide
for sh in wb.worksheets:
    sh.sheet_properties.tabColor = {'Dashboard': '0E3A40', 'Calls': 'B4541A', 'Projects': '1C5F68'}.get(sh.title, '9AA5A6')
ws.sheet_view.zoomScale = 90
wp.sheet_view.zoomScale = 90
wb.active = wb.sheetnames.index('Calls')
for sh in wb.worksheets:
    sh.sheet_view.tabSelected = sh.title == 'Calls'
ws.print_title_rows = '1:1'
for sh in (wd, wm):
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT)
print('rows', LAST - 1, 'projects', PLAST - 1)
