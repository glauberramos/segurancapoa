"""Gera municipios.js a partir das planilhas da SSP-RS na raiz do repositório.

Uso: python3 scripts/build_municipios.py
Requer: openpyxl, xlrd
"""
import glob
import json
import os
import re
import unicodedata

import openpyxl
import xlrd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MONTHS = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']

# Região Metropolitana de Porto Alegre (34 municípios) — população: IBGE, estimativa 2025.
MUNICIPIOS = {
    'PORTO ALEGRE': ('Porto Alegre', 1388794),
    'ALVORADA': ('Alvorada', 194062),
    'ARARICA': ('Araricá', 8831),
    'ARROIO DOS RATOS': ('Arroio dos Ratos', 14934),
    'CACHOEIRINHA': ('Cachoeirinha', 141503),
    'CAMPO BOM': ('Campo Bom', 64719),
    'CANOAS': ('Canoas', 359840),
    'CAPELA DE SANTANA': ('Capela de Santana', 11424),
    'CHARQUEADAS': ('Charqueadas', 36110),
    'DOIS IRMAOS': ('Dois Irmãos', 31804),
    'ELDORADO DO SUL': ('Eldorado do Sul', 41012),
    'ESTANCIA VELHA': ('Estância Velha', 49499),
    'ESTEIO': ('Esteio', 78144),
    'GLORINHA': ('Glorinha', 7851),
    'GRAVATAI': ('Gravataí', 275430),
    'GUAIBA': ('Guaíba', 95946),
    'IGREJINHA': ('Igrejinha', 34776),
    'IVOTI': ('Ivoti', 23566),
    'MONTENEGRO': ('Montenegro', 66352),
    'NOVA HARTZ': ('Nova Hartz', 20558),
    'NOVA SANTA RITA': ('Nova Santa Rita', 30189),
    'NOVO HAMBURGO': ('Novo Hamburgo', 235802),
    'PAROBE': ('Parobé', 53569),
    'PORTAO': ('Portão', 35273),
    'ROLANTE': ('Rolante', 21752),
    'SANTO ANTONIO DA PATRULHA': ('Santo Antônio da Patrulha', 44431),
    'SAO JERONIMO': ('São Jerônimo', 21412),
    'SAO LEOPOLDO': ('São Leopoldo', 225737),
    'SAO SEBASTIAO DO CAI': ('São Sebastião do Caí', 24898),
    'SAPIRANGA': ('Sapiranga', 77935),
    'SAPUCAIA DO SUL': ('Sapucaia do Sul', 136572),
    'TAQUARA': ('Taquara', 54353),
    'TRIUNFO': ('Triunfo', 28435),
    'VIAMAO': ('Viamão', 231996),
}


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().upper()


def col_key(h):
    h = norm(h)
    if h == 'HOMICIDIO DOLOSO':
        return 'hom'
    if h == 'TOTAL DE VITIMAS DE HOMICIDIO DOLOSO':
        return 'vit_hom'
    if h == 'LATROCINIO':
        return 'latro'
    if h == 'VITIMAS DE LATROCINIO':
        return 'vit_latro'
    if h == 'VITIMAS DE LESAO CORP. SEG. MORTE':
        return 'vit_lesao'
    if 'CVLI' in h:
        return 'cvli'
    if h == 'ROUBOS':
        return 'roubos'
    if h == 'ROUBO DE VEICULO':
        return 'roubo_veic'
    if h == 'ESTELIONATO':
        return 'estel'
    return None


def parse_rows(rows):
    """rows: lista de listas de células de uma aba mensal. Retorna {MUNICIPIO: {chave: valor}}."""
    header = None
    out = {}
    for r in rows:
        first = norm(r[0]) if r else ''
        if header is None:
            if first in ('MUNICIPIOS', 'MUNICIPIO', 'NOME DA CIDADE'):
                header = [col_key(c) for c in r]
            continue
        if first in MUNICIPIOS:
            out[first] = {k: int(v or 0) for k, v in zip(header, r) if k}
    return out


def cvli(d, year):
    if year == 2011:  # só ocorrências disponíveis
        return d['hom'] + d['latro']
    if year >= 2018:  # total oficial da SSP-RS (inclui feminicídio e mortes por intervenção policial)
        return d['cvli']
    # 2012–2017: as planilhas não trazem o total de CVLI
    return d['vit_hom'] + d['vit_latro'] + d['vit_lesao']


def find(pattern):
    files = sorted(glob.glob(os.path.join(ROOT, pattern)))
    assert files, pattern
    return files[-1]


def load_year(year):
    if year == 2011:
        wb = xlrd.open_workbook(find('*por-municipio-2011-por-mes.xls'))
        return [parse_rows([ws.row_values(i) for i in range(ws.nrows)])
                for ws in (wb.sheet_by_name(m) for m in MONTHS)]
    if year == 2026:
        path = find('*-jul-2026-*.xlsx')
    else:
        path = [f for f in glob.glob(os.path.join(ROOT, '*.xlsx'))
                if re.search(rf'(?<!\d){year}(?!\d)', os.path.basename(f).split('atualizado')[0])]
        assert len(path) == 1, (year, path)
        path = path[0]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = [w.title for w in wb.worksheets]
    return [parse_rows([list(r) for r in wb[m].iter_rows(values_only=True)])
            for m in MONTHS if m in names]


def main():
    series = {k: {'cvli': [], 'roubos': [], 'roubo_veic': [], 'estel': []} for k in MUNICIPIOS}
    labels = []
    for year in range(2011, 2027):
        for i, month in enumerate(load_year(year)):
            labels.append(f"{MONTHS[i].capitalize()}/{str(year)[2:]}")
            for k in MUNICIPIOS:
                d = month[k]
                s = series[k]
                s['cvli'].append(cvli(d, year))
                s['roubos'].append(d['roubos'])
                s['roubo_veic'].append(d['roubo_veic'])
                s['estel'].append(d['estel'])

    data = {
        'labels': labels,
        'municipios': {
            k: {'nome': MUNICIPIOS[k][0], 'populacao': MUNICIPIOS[k][1], **series[k]} for k in MUNICIPIOS
        },
    }
    with open(os.path.join(ROOT, 'municipios.js'), 'w', encoding='utf-8') as f:
        f.write('// Gerado por scripts/build_municipios.py a partir das planilhas da SSP-RS. Não editar à mão.\n')
        f.write('window.MUNICIPIOS_DATA = ')
        f.write(json.dumps(data, ensure_ascii=False, separators=(',', ':')))
        f.write(';\n')
    print(len(labels), 'meses,', len(MUNICIPIOS), 'municípios')


if __name__ == '__main__':
    main()
