# Valida que las partes modificadas/nuevas del .xlsx sean XML bien formado y coherentes
import sys, zipfile, re
from xml.dom import minidom
z = zipfile.ZipFile(sys.argv[1])
parts = ['xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml','xl/styles.xml'] + [n for n in z.namelist() if 'sheetReorg' in n]
for p in parts:
    minidom.parseString(z.read(p))
wb = z.read('xl/workbook.xml').decode()
st = z.read('xl/styles.xml').decode()
nx = len(re.findall(r'<xf[\s>/]', re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', st, re.S).group(1)))
cnt = int(re.search(r'<cellXfs count="(\d+)"', st).group(1))
maxs = max(int(x) for p in parts if 'sheetReorg' in p for x in re.findall(r' s="(\d+)"', z.read(p).decode()))
print('  XML OK |', 'cellXfs', nx, 'count', cnt, 'max s usado', maxs, '| calcPr', re.search(r'<calcPr[^>]*>', wb).group(0), '|', re.findall(r'<definedName name="_xlnm._FilterDatabase"[^>]*>', wb)[-1])
assert nx == cnt and maxs < nx
