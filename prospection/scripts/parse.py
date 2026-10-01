"""Parse Infopro/Baublatt 'DRINGEND' project PDFs (one project per page)."""
import glob, json, os, re, sys
import pdfplumber

SRC = sys.argv[1] if len(sys.argv) > 1 else '.'
COL = 290  # x split between left/right columns


def nat_key(name):
    parts = re.split(r'(\d+)', name)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def lines_of(words, tol=3):
    """Group words into lines by 'top'."""
    out = []
    for w in sorted(words, key=lambda w: (w['top'], w['x0'])):
        if out and abs(out[-1]['top'] - w['top']) <= tol:
            out[-1]['words'].append(w)
        else:
            out.append({'top': w['top'], 'words': [w]})
    for l in out:
        l['words'].sort(key=lambda w: w['x0'])
    return out


def txt(ws):
    return ' '.join(w['text'] for w in ws).strip()


def bold(w):
    return 'Bold' in w['fontname']


def find_line(lines, pred):
    for i, l in enumerate(lines):
        if pred(l):
            return i
    return None


def starts(l, text, x_max=60):
    t = txt(l['words'])
    return t.startswith(text) and l['words'][0]['x0'] < x_max


def parse_page(page):
    words = page.extract_words(extra_attrs=['fontname', 'size'])
    L = lines_of(words)
    i_termine = find_line(L, lambda l: starts(l, 'Termine und Fakten'))
    i_allg = find_line(L, lambda l: starts(l, 'Allgemein'))
    i_ukat = find_line(L, lambda l: starts(l, 'Unterkategorien'))
    i_kont = find_line(L, lambda l: starts(l, 'Alle wichtigen Kontakte'))
    i_hinw = find_line(L, lambda l: starts(l, 'Weitere Hinweise'))
    i_copy = find_line(L, lambda l: 'Copyright' in txt(l['words']))
    n = len(L)
    end_kont = i_hinw if i_hinw is not None else (i_copy if i_copy is not None else n)

    p = {}
    # ---------- header ----------
    head = L[:i_termine]
    p['lieu_titre'] = txt([w for w in head[0]['words'] if w['x0'] < 400])
    title_parts, stage_parts = [], []
    for l in head[1:]:
        left = [w for w in l['words'] if w['x0'] < 420]
        right = [w for w in l['words'] if w['x0'] >= 420]
        if txt(left) and txt(left) != 'Baustadium':
            title_parts.append(txt(left))
        if right and txt(right) != 'Baustadium':
            stage_parts.append(txt(right))
    p['projet'] = ' '.join(title_parts)
    p['stade'] = ' '.join(stage_parts)

    # ---------- Lage / Termine ----------
    lage, termine = [], []
    for l in L[i_termine + 2:i_allg]:
        left = [w for w in l['words'] if w['x0'] < COL]
        right = [w for w in l['words'] if w['x0'] >= COL]
        if left and txt(left) != 'Lage':
            lage.append(txt(left))
        if right and txt(right) != 'Termine':
            termine.append(txt(right))
    p['reference'] = next((x.split(':', 1)[1].strip() for x in lage if x.startswith('Referenz')), '')
    p['adresse_chantier'] = ', '.join(x for x in lage if not x.startswith('Referenz'))
    p['termine'] = []
    for t in termine:
        m = re.match(r'(.*?)\s+(\d{2}\.\d{2}\.\d{4})$', t)
        if m:
            p['termine'].append({'evenement': m.group(1), 'date': m.group(2)})

    # ---------- Allgemein / Kennzahlen ----------
    kv = {}
    for l in L[i_allg + 1:i_ukat]:
        for side in ([w for w in l['words'] if w['x0'] < COL], [w for w in l['words'] if w['x0'] >= COL]):
            if not side:
                continue
            # label = words before the value column (value starts ~95px after label start)
            x_lab = side[0]['x0']
            lab = [w for w in side if w['x0'] < x_lab + 90]
            val = [w for w in side if w['x0'] >= x_lab + 90]
            if lab:
                kv[txt(lab)] = txt(val)
    p['kv'] = kv

    # ---------- Unterkategorien ----------
    ucols = {'L': [], 'R': []}
    for l in L[i_ukat + 1:i_kont]:
        for key, side in (('L', [w for w in l['words'] if w['x0'] < 295]),
                          ('R', [w for w in l['words'] if w['x0'] >= 295])):
            if not side:
                continue
            first = side[0]['text']
            if re.fullmatch(r'\d{2,4}', first) and side[0]['x0'] < (60 if key == 'L' else 320):
                ucols[key].append(txt(side))
            elif ucols[key]:
                ucols[key][-1] += ' ' + txt(side)
            else:
                ucols[key].append(txt(side))
    p['sous_categories'] = ucols['L'] + ucols['R']

    # ---------- Kontakte ----------
    kl = L[i_kont + 1:end_kont]
    contacts, extra = [], []
    # compact table: lines whose first word at x<40 regular font and that have a word >= x390 numeric (PLZ)
    def is_compact(l):
        ws = l['words']
        return (ws[0]['x0'] < 40 and not bold(ws[0]) and any(110 < w['x0'] < 135 for w in ws)
                and any(380 < w['x0'] < 400 and re.fullmatch(r'\d{4}', w['text']) for w in ws))
    i_compact = next((i for i, l in enumerate(kl) if is_compact(l)), None)
    grid = kl if i_compact is None else kl[:i_compact]
    # drop lone numeric bold marker lines (e.g. "22")
    grid = [l for l in grid if not re.fullmatch(r'\d+', txt(l['words']))]
    ROLES = ('Bauherr', 'Architekt/Planer')
    # block starts: lines containing a bold size>=10.5 role header
    for side_lo, side_hi in ((0, COL), (COL, 9999)):
        cur = None
        for l in grid:
            ws = [w for w in l['words'] if side_lo <= w['x0'] < side_hi]
            if not ws:
                continue
            t = txt(ws)
            if bold(ws[0]) and ws[0]['size'] > 10.5:
                cur = {'role': t, 'nom': [], 'lignes': [], 'top': l['top'], 'col': 'L' if side_lo == 0 else 'R'}
                contacts.append(cur)
                continue
            if cur is None:
                continue
            if bold(ws[0]) and ws[0]['text'] != 'Kontakt:' and not cur['lignes']:
                cur['nom'].append(t)
            else:
                cur['lignes'].append(t)
    contacts.sort(key=lambda c: (c['top'], c['col']))
    out = []
    for c in contacts:
        d = {'role': c['role'], 'societe': ' '.join(c['nom']), 'tel': '', 'fax': '', 'email': '',
             'web': '', 'interlocuteurs': [], 'adresse': [], 'npa': '', 'localite': '', 'source': 'fiche'}
        for t in c['lignes']:
            if t.startswith('Tel.:'):
                d['tel'] = t[5:].strip()
            elif t.startswith('Fax:'):
                d['fax'] = t[4:].strip()
            elif t.startswith('Mobile:') or t.startswith('Natel:'):
                d['tel'] = d['tel'] or t.split(':', 1)[1].strip()
            elif '@' in t and ' ' not in t:
                d['email'] = t
            elif t.startswith('http') or t.startswith('www.'):
                d['web'] = t
            elif t.startswith('Kontakt:'):
                d['interlocuteurs'].append(t[8:].strip())
            else:
                d['adresse'].append(t)
        for a in d['adresse']:
            m = re.match(r'^(?:CH-)?(\d{4})\s+(.+)$', a)
            if m:
                d['npa'], d['localite'] = m.group(1), m.group(2)
        out.append(d)
    # compact table rows — read chars in drawing (stream) order: the PDF can print a
    # wrapped name line almost on top of the next row, which garbles position-sorted text.
    if i_compact is not None:
        y0 = kl[i_compact]['top'] - 2
        y1 = L[end_kont]['top'] - 2 if end_kont < n else page.height
        chars = [c for c in page.chars if y0 <= c['top'] < y1]
        runs = []
        for c in chars:
            r = runs[-1] if runs else None
            if (r is None or abs(c['top'] - r['top']) > 0.05 or c['x0'] < r['x1'] - 1
                    or c['x0'] - r['x1'] > 6):
                runs.append({'top': c['top'], 'x0': c['x0'], 'x1': c['x1'], 't': c['text']})
            else:
                r['t'] += c['text']
                r['x1'] = c['x1']
        cur = None
        for r in runs:
            t = r['t'].strip()
            if not t:
                continue
            if r['x0'] < 40:
                cur = {'role': t, 'societe': '', 'city': '', 'tel': '', 'top': r['top']}
                extra.append(cur)
            elif cur is None:
                continue
            elif r['x0'] < 385:
                if not cur['societe']:
                    sep = ''
                elif abs(r['top'] - cur.get('_ltop', r['top'])) > 0.05:
                    sep = ' - '  # wrapped onto a new line
                elif r['x0'] - cur.get('_lx1', 0) < 2:
                    sep = ''  # same word split into two runs (kerning)
                else:
                    sep = ' '
                cur['societe'] = (cur['societe'] + sep + t).strip()
                cur['_ltop'], cur['_lx1'] = r['top'], r['x1']
            elif r['x0'] < 485:
                cur['city'] = (cur['city'] + ' ' + t).strip()
            else:
                cur['tel'] = (cur['tel'] + ' ' + t).strip()
        for e in extra:
            m = re.match(r'^(\d{4})\s+(.+)$', e.pop('city'))
            for k in ('top', '_ltop', '_lx1'):
                e.pop(k, None)
            e.update({'fax': '', 'email': '', 'web': '', 'interlocuteurs': [], 'adresse': [],
                      'npa': m.group(1) if m else '', 'localite': m.group(2) if m else '',
                      'source': 'liste complémentaire'})
    p['contacts'] = out + extra

    # ---------- Weitere Hinweise ----------
    if i_hinw is not None:
        stop = i_copy if i_copy is not None else n
        p['description'] = ' '.join(txt(l['words']) for l in L[i_hinw + 1:stop])
    else:
        p['description'] = ''
    return p


def main():
    files = sorted(glob.glob(os.path.join(SRC, '*.pdf')), key=lambda f: nat_key(os.path.basename(f)))
    res = []
    for f in files:
        with pdfplumber.open(f) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                try:
                    p = parse_page(page)
                except Exception as e:  # keep going, report
                    p = {'erreur': repr(e)}
                p['fichier'] = os.path.basename(f)
                p['page'] = i
                p['pages_total'] = len(pdf.pages)
                res.append(p)
    json.dump(res, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
