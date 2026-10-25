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

# População: IBGE, estimativa 2025.
# Região Metropolitana de Porto Alegre (34 municípios)
RMPA = {
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

# Aglomeração Urbana do Litoral Norte (20 municípios)
LITORAL_NORTE = {
    'ARROIO DO SAL': ('Arroio do Sal', 11418),
    'BALNEARIO PINHAL': ('Balneário Pinhal', 15413),
    'CAPAO DA CANOA': ('Capão da Canoa', 66234),
    'CAPIVARI DO SUL': ('Capivari do Sul', 4080),
    'CARAA': ('Caraá', 7555),
    'CIDREIRA': ('Cidreira', 17583),
    'DOM PEDRO DE ALCANTARA': ('Dom Pedro de Alcântara', 2617),
    'IMBE': ('Imbé', 28027),
    'ITATI': ('Itati', 2696),
    'MAMPITUBA': ('Mampituba', 3203),
    'MAQUINE': ('Maquiné', 7593),
    'MORRINHOS DO SUL': ('Morrinhos do Sul', 3133),
    'OSORIO': ('Osório', 48999),
    'PALMARES DO SUL': ('Palmares do Sul', 13189),
    'TERRA DE AREIA': ('Terra de Areia', 10575),
    'TORRES': ('Torres', 43344),
    'TRAMANDAI': ('Tramandaí', 56430),
    'TRES CACHOEIRAS': ('Três Cachoeiras', 11224),
    'TRES FORQUILHAS': ('Três Forquilhas', 2813),
    'XANGRI-LA': ('Xangri-lá', 16949),
}

# Serra: Região Metropolitana da Serra Gaúcha (13 municípios) + Gramado, Canela,
# Nova Petrópolis e São Francisco de Paula (Hortênsias)
SERRA = {
    'ANTONIO PRADO': ('Antônio Prado', 13332),
    'BENTO GONCALVES': ('Bento Gonçalves', 127977),
    'CARLOS BARBOSA': ('Carlos Barbosa', 31585),
    'CAXIAS DO SUL': ('Caxias do Sul', 479599),
    'FARROUPILHA': ('Farroupilha', 72552),
    'FLORES DA CUNHA': ('Flores da Cunha', 32015),
    'GARIBALDI': ('Garibaldi', 35563),
    'IPE': ('Ipê', 5490),
    'MONTE BELO DO SUL': ('Monte Belo do Sul', 2608),
    'NOVA PADUA': ('Nova Pádua', 2389),
    'PINTO BANDEIRA': ('Pinto Bandeira', 2785),
    'SANTA TEREZA': ('Santa Tereza', 1529),
    'SAO MARCOS': ('São Marcos', 21549),
    'CANELA': ('Canela', 50715),
    'GRAMADO': ('Gramado', 41705),
    'NOVA PETROPOLIS': ('Nova Petrópolis', 23934),
    'SAO FRANCISCO DE PAULA': ('São Francisco de Paula', 22388),
}

MUNICIPIOS = {
    **{k: (*v, 'Região Metropolitana') for k, v in RMPA.items()},
    **{k: (*v, 'Litoral Norte') for k, v in LITORAL_NORTE.items()},
    **{k: (*v, 'Serra') for k, v in SERRA.items()},
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
                d = month.get(k)  # None antes da emancipação (ex.: Pinto Bandeira em 2011)
                s = series[k]
                s['cvli'].append(cvli(d, year) if d else None)
                s['roubos'].append(d['roubos'] if d else None)
                s['roubo_veic'].append(d['roubo_veic'] if d else None)
                s['estel'].append(d['estel'] if d else None)

    data = {
        'labels': labels,
        'municipios': {
            k: {'nome': MUNICIPIOS[k][0], 'populacao': MUNICIPIOS[k][1], 'regiao': MUNICIPIOS[k][2], **series[k]}
            for k in MUNICIPIOS
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
