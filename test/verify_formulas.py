"""Evalúa las fórmulas de las hojas nuevas (VINCULADO) usando los valores originales
cacheados por Excel en el libro fuente, y compara con lo esperado.
uso: python3 verify_formulas.py ORIGINAL.xlsx VINCULADO.xlsx
"""
import sys, re, zipfile
import openpyxl
from openpyxl.utils import column_index_from_string as cidx

orig_path, out_path = sys.argv[1], sys.argv[2]
wbf = openpyxl.load_workbook(out_path, read_only=True, data_only=False)
new_names = [n for n in wbf.sheetnames if re.match(r'^(PPTO x CAPITULO|VERIFICACION|MATRIZ CAP x UE)', n)]
sheets = {}
for n in new_names:
    d = {}
    for row in wbf[n].iter_rows():
        for c in row:
            if c.value is not None and hasattr(c, 'column'):
                d[(c.row, c.column)] = c.value
    sheets[n] = d

# hojas fuente referenciadas
src_refs = set()
for d in sheets.values():
    for v in d.values():
        if isinstance(v, str) and v.startswith('='):
            for m in re.finditer(r"'((?:[^']|'')+)'!", v):
                nm = m.group(1).replace("''", "'")
                if nm not in sheets: src_refs.add(nm)
wbo = openpyxl.load_workbook(orig_path, read_only=True, data_only=True)
src = {}
for nm in src_refs:
    d = {}
    for row in wbo[nm].iter_rows():
        for c in row:
            if c.value is not None and hasattr(c, 'column'):
                d[(c.row, c.column)] = c.value
    src[nm] = d

TOK = re.compile(r"""\s*(?:
 (?P<ref>(?:'(?:[^']|'')+'!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|
 (?P<colrange>(?:'(?:[^']|'')+'!)?\$?[A-Z]{1,3}:\$?[A-Z]{1,3})|
 (?P<num>\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)|
 (?P<str>"(?:[^"]|"")*")|
 (?P<func>[A-Z]+)\(|
 (?P<op><=|>=|<>|[-+*/(),<>=&]))""", re.X)

memo = {}
def cell_value(sheet, r, c):
    if sheet in sheets:
        key = (sheet, r, c)
        if key in memo: return memo[key]
        v = sheets[sheet].get((r, c))
        if isinstance(v, str) and v.startswith('='):
            memo[key] = None
            v = evaluate(v[1:], sheet)
        memo[key] = v
        return v
    return src[sheet].get((r, c))

def split_sheet(tok, cur):
    if '!' in tok:
        s, a = tok.rsplit('!', 1)
        return s[1:-1].replace("''", "'"), a
    return cur, tok

def parse_cell(a):
    m = re.match(r'\$?([A-Z]+)\$?(\d+)', a); return int(m.group(2)), cidx(m.group(1))

def evaluate(f, cur):
    toks = []
    pos = 0
    while pos < len(f):
        m = TOK.match(f, pos)
        if not m: raise ValueError('token? ' + f[pos:pos+30])
        pos = m.end()
        kind = m.lastgroup; toks.append((kind, m.group(kind)))
    i = 0
    def peek(): return toks[i] if i < len(toks) else (None, None)
    def nxt():
        nonlocal i; t = toks[i]; i += 1; return t
    def expr():
        v = comp()
        return v
    def comp():
        a = add()
        while peek()[1] in ('<', '>', '=', '<=', '>=', '<>'):
            op = nxt()[1]; b = add()
            a = {'<': a < b, '>': a > b, '=': a == b, '<=': a <= b, '>=': a >= b, '<>': a != b}[op]
        return a
    def add():
        a = mul()
        while peek()[1] in ('+', '-'):
            op = nxt()[1]; b = mul(); a = num(a) + num(b) if op == '+' else num(a) - num(b)
        return a
    def mul():
        a = unary()
        while peek()[1] in ('*', '/'):
            op = nxt()[1]; b = unary(); a = num(a) * num(b) if op == '*' else num(a) / num(b)
        return a
    def unary():
        if peek()[1] == '-': nxt(); return -num(unary())
        if peek()[1] == '+': nxt(); return num(unary())
        return atom()
    def atom():
        k, t = nxt()
        if k == 'num': return float(t)
        if k == 'str': return t[1:-1].replace('""', '"')
        if k == 'ref':
            s, a = split_sheet(t, cur)
            if ':' in a:
                p, q = a.split(':'); r1, c1 = parse_cell(p); r2, c2 = parse_cell(q)
                return ('range', s, r1, c1, r2, c2)
            r, c = parse_cell(a); return cell_value(s, r, c)
        if k == 'colrange':
            s, a = split_sheet(t, cur); p, q = a.replace('$', '').split(':')
            return ('col', s, cidx(p))
        if k == 'op' and t == '(':
            v = expr(); nxt(); return v
        if k == 'func':
            args = []
            if peek()[1] != ')':
                args.append(expr())
                while peek()[1] == ',': nxt(); args.append(expr())
            nxt()
            return call(t, args)
        raise ValueError(f'atom {k} {t}')
    v = expr()
    return v

def num(v):
    if isinstance(v, bool): return 1.0 if v else 0.0
    if isinstance(v, (int, float)): return float(v)
    return 0.0

def values(arg):
    if isinstance(arg, tuple) and arg[0] == 'range':
        _, s, r1, c1, r2, c2 = arg
        return [cell_value(s, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
    if isinstance(arg, tuple) and arg[0] == 'col':
        _, s, c = arg
        rows = max(r for (r, cc) in sheets[s]) if s in sheets else 0
        return [cell_value(s, r, c) for r in range(1, rows + 1)]
    return [arg]

def call(fn, args):
    if fn == 'SUM': return sum(num(x) for a in args for x in values(a) if isinstance(x, (int, float)) and not isinstance(x, bool))
    if fn == 'ABS': return abs(num(args[0]))
    if fn == 'ROUND': return round(num(args[0]), int(num(args[1])))
    if fn == 'AND': return all(bool(a) for a in args)
    if fn == 'IF': return args[1] if args[0] else (args[2] if len(args) > 2 else False)
    if fn == 'HYPERLINK': return args[1]
    if fn == 'SUMIFS':
        sv = values(args[0]); tot = 0.0
        crits = [(values(args[k]), str(args[k + 1])) for k in range(1, len(args), 2)]
        for idx, x in enumerate(sv):
            ok = True
            for cv, cr in crits:
                target = re.sub(r'~([~*?])', r'\1', cr[1:] if cr.startswith('=') else cr)
                val = cv[idx] if idx < len(cv) else None
                if not isinstance(val, str) or val.strip().upper() != target.strip().upper(): ok = False; break
            if ok and isinstance(x, (int, float)): tot += x
        return tot
    raise ValueError('función no soportada ' + fn)

# evaluar todas las fórmulas
errs = 0; nform = 0
for n in new_names:
    for (r, c), v in sheets[n].items():
        if isinstance(v, str) and v.startswith('='):
            nform += 1
            try: cell_value(n, r, c)
            except Exception as e:
                errs += 1
                if errs < 5: print('  ERROR', n, r, c, v[:80], e)
chk = [n for n in new_names if n.startswith('VERIFICACION')][0]
states = {}
for (r, c), v in sorted(sheets[chk].items()):
    val = cell_value(chk, r, c)
    if val in ('OK', 'REVISAR'):
        states[val] = states.get(val, 0) + 1
        if val == 'REVISAR':
            print('  REVISAR fila', r, [cell_value(chk, r, k) for k in range(1, 6)])
main = [n for n in new_names if n.startswith('PPTO x CAPITULO')][0]
tot = [(r, c) for (r, c), v in sheets[main].items() if v == 'TOTAL PRESUPUESTO'][0]
vt_col = [c for (r, c), v in sheets[main].items() if r == 6 or r == 5 or r == 7]
print(f'  fórmulas evaluadas={nform} errores={errs} estados={states}')
mat = [n for n in new_names if n.startswith('MATRIZ')][0]
for (r, c), v in sheets[mat].items():
    if isinstance(v, str) and v.startswith('Control'):
        cc = max(k for (rr, k) in sheets[mat] if rr == r)
        print('  matriz control =', cell_value(mat, r, cc))
