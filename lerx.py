# -*- coding: utf-8 -*-
"""Leitor minimo de xlsx via XML, para arquivos que o openpyxl recusa."""
import zipfile, re
import xml.etree.ElementTree as ET
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

class Livro:
    def __init__(self, caminho):
        self.z = zipfile.ZipFile(caminho)
        wb = ET.fromstring(self.z.read('xl/workbook.xml'))
        rels = ET.fromstring(self.z.read('xl/_rels/workbook.xml.rels'))
        alvo = {r.get('Id'): r.get('Target') for r in rels}
        self.abas = []
        for s in wb.find(NS+'sheets'):
            t = alvo[s.get(R+'id')]
            self.abas.append((s.get('name'), 'xl/' + t.lstrip('/')))
        self.ss = []
        if 'xl/sharedStrings.xml' in self.z.namelist():
            for si in ET.fromstring(self.z.read('xl/sharedStrings.xml')):
                self.ss.append(''.join(t.text or '' for t in si.iter(NS+'t')))
    def celulas(self, arquivo):
        """devolve {(linha, coluna_indice0): valor}"""
        out = {}
        raiz = ET.fromstring(self.z.read(arquivo))
        for row in raiz.iter(NS+'row'):
            n = int(row.get('r'))
            for c in row:
                ref = c.get('r')
                col = re.match(r'([A-Z]+)', ref).group(1)
                idx = 0
                for ch in col: idx = idx*26 + (ord(ch)-64)
                v = c.find(NS+'v'); t = c.get('t')
                if t == 'inlineStr':
                    val = ''.join(x.text or '' for x in c.iter(NS+'t'))
                elif v is None: val = None
                elif t == 's':  val = self.ss[int(v.text)]
                else:           val = v.text
                if val not in (None, ''): out[(n, idx-1)] = val
        return out
