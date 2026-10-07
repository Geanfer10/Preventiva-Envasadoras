"""Leitor da Programação de Produção Diária (PDF do PCP) -> JSON para o painel de preventiva.

Uso: python parse_prog.py PROGRAMACAO.pdf saida.json
Lê só os produtos acabados (códigos 111xxxx), atribui horários e status a cada turno
pela posição na página e calcula as janelas (horas sem produção) por máquina.
"""
import re, sys, json, unicodedata
import pdfplumber

# turnos padrão, em minutos contados a partir de 06:30
TURNOS = {1: (0, 485), 2: (485, 965), 3: (965, 1440)}   # 06:30-14:35, 14:35-22:35, 22:35-06:30
COLS = {1: (440, 534), 2: (534, 618), 3: (618, 705)}      # faixa x de cada coluna de turno no PDF
ST_OK = {'PROD', 'CIP', 'SETUP', 'ENXAGUE', 'LIMPEZA', 'TESTE', 'PARADA'}
# duração padrão (min) de eventos que o PDF marca sem horário próprio; o evento é posto no início do turno
DUR_EVENTO = {'ENXAGUE': 40}
MESES = {m: i + 1 for i, m in enumerate(['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho',
                                           'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'])}
MP_LETRA = {'A': 'ATO-MP01', 'B': 'ATO-MP02', 'C': 'ATO-MP03', 'D': 'ATO-MP04', 'E': 'ATO-MP05', 'F': 'ATO-MP07',
            'R1': 'VEG-MP06', 'R2': 'VEG-MP08'}


def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper()


def maquinas_da_linha(linha):
    """Converte o texto da coluna LINHA nos ids de máquina do painel."""
    L = sem_acento(linha)
    base = L.split(' - ')[0]                      # '60L 1 & 2 - 170G' -> '60L 1 & 2'
    out = []
    if 'RETORT' in L:
        out += [MP_LETRA['R' + n] for n in re.findall(r'\b([12])\b', base)] or [MP_LETRA['R1'], MP_LETRA['R2']]
    elif 'MESPACK' in L:
        for tk in re.findall(r'\b(R[12]|[A-F])\b', base.replace('MESPACK', '')):
            out.append(MP_LETRA[tk])
    elif re.search(r'\b60\s*L', L):
        out += [{'1': 'VEG-EC02', '2': 'VEG-EC03'}[n] for n in re.findall(r'\b([12])\b', base.replace('60L', ''))]
    elif 'SERAC' in L or 'TOP DOWN' in L or 'TOPDOWN' in L or 'FRASCO' in L:
        out.append('ATO-EC08')
    elif 'CONDOR' in L or '1,7' in L or '1.7' in L:
        out.append('VEG-RC07')
    elif 'COPO' in L or 'LATA' in L:
        out.append('ATO-EC06')
    elif '3100' in L or '4KG' in L:
        out.append('ATO-EC07')
    elif 'COMACO' in L:
        out.append('VEG-EC01')
    elif 'POTE' in L:
        out.append('VEG-EC04')
    return out


def hm(s):
    h, m = map(int, s.split(':'))
    return (h * 60 + m - 390) % 1440               # minutos desde 06:30


def ler(path):
    """Lê todas as páginas; cada página com data própria vira uma programação."""
    out = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            r = ler_pagina(pg.extract_words(), pg.extract_text() or '', pg.chars)
            if r['data']:
                out.append(r)
    return out


def _colunas(W):
    """Acha as colunas pelo cabeçalho da própria página (o PCP muda a largura das colunas de um dia para o outro)."""
    prog = sorted({round(w['x0']) for w in W if w['text'] == 'PROG.'})
    if len(prog) < 3:
        return COLS, 350, 425
    # o status (CIP, ENXAGUE...) começa um pouco antes do título "PROG."; a 3ª coluna tem a largura da 2ª
    m = 18
    cols = {1: (prog[0] - m, prog[1] - m), 2: (prog[1] - m, prog[2] - m), 3: (prog[2] - m, 2 * prog[2] - prog[1] - m)}
    lx = [w['x0'] for w in W if w['text'] == 'LINHA']
    um = [w['x0'] for w in W if w['text'] == 'UM' and w['x0'] < prog[0]]
    lin_ini = (min(lx) - 25) if lx else 345
    lin_fim = (min(um) - 2) if um else prog[0] - 15
    return cols, lin_ini, lin_fim


def _linhas(cs):
    """Agrupa caracteres pela altura (cada linha de texto separada) e devolve o texto de cada linha."""
    L = []
    for c in sorted(cs, key=lambda c: c['top']):
        if L and abs(c['top'] - L[-1][0]) < 1.2:
            L[-1][1].append(c)
        else:
            L.append([c['top'], [c]])
    return [''.join(c['text'] for c in sorted(g, key=lambda c: c['x0'])) for _, g in L]


def ler_pagina(W, texto, C=None):
    # data da programação e emissão
    md = re.search(r'(\d{1,2}) de (\w+) de (\d{4})', texto)
    data = f"{md.group(3)}-{MESES[md.group(2).lower()]:02d}-{int(md.group(1)):02d}" if md else None
    me = re.search(r'(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})', texto)
    emissao = f"{me.group(1)} {me.group(2)}" if me else ''
    obs = []
    if 'OBSERVAÇÕES IMPORTANTES' in texto:
        obs = [l.strip() for l in texto.split('OBSERVAÇÕES IMPORTANTES', 1)[1].splitlines() if l.strip()]
    cols, lin_ini, lin_fim = _colunas(W)
    C = C or []

    prods = sorted([w for w in W if re.fullmatch(r'111\d{4}', w['text']) and w['x0'] < 70], key=lambda w: w['top'])
    linhas = []
    for i, pw in enumerate(prods):
        y = pw['top']
        y0 = (prods[i - 1]['top'] + 7) if i else y - 25
        nxt = prods[i + 1]['top'] if i + 1 < len(prods) else y + 30
        y1 = min(y + 7, nxt - 11)        # o status fica um pouco abaixo do código; os horários da próxima linha ficam ~6 acima dela
        band = [w for w in W if y0 < w['top'] <= y1]
        desc = ' '.join(w['text'] for w in sorted(band, key=lambda w: w['x0']) if abs(w['top'] - y) < 2 and 75 < w['x0'] < lin_ini)
        linha = ' '.join(w['text'] for w in sorted(band, key=lambda w: w['x0'])
                         if y - 12 < w['top'] < y + 2 and lin_ini <= w['x0'] < lin_fim)
        bandc = [c for c in C if y0 < c['top'] <= y1 and c['text'].strip()]
        turnos = []
        for t, (xa, xb) in cols.items():
            cc = [c for c in bandc if xa <= c['x0'] < xb]
            # o PDF sobrepõe três textos na mesma célula; o tamanho da letra separa cada um:
            # pequeno = horário início/fim, médio = status (PROD/CIP/ENXAGUE...), grande = quantidade
            horas = [h for ln in _linhas([c for c in cc if c['size'] < 5.0]) if not re.search(r'[A-Za-z]', ln)
                     for h in re.findall(r'\d{2}:\d{2}', ln)]
            st = next((sem_acento(ln) for ln in _linhas([c for c in cc if 5.0 <= c['size'] < 9])
                       if sem_acento(ln) in ST_OK), None)
            nums = ''.join(ln for ln in _linhas([c for c in cc if c['size'] >= 9]) if re.fullmatch(r'[\d.,\s-]+', ln))
            qtd = int(re.sub(r'[^\d]', '', nums)) if re.search(r'\d', nums) else 0
            if len(horas) >= 2 or st or qtd:
                ini = hm(horas[0]) if len(horas) >= 2 else None
                fim = hm(horas[1]) if len(horas) >= 2 else None
                if ini is not None and fim is not None and fim <= ini:
                    fim += 1440
                turnos.append(dict(t=t, ini=ini, fim=fim, h=f"{horas[0]}–{horas[1]}" if len(horas) >= 2 else '',
                                   st=st or 'PROD', qtd=qtd))
        linhas.append(dict(cod=pw['text'], desc=desc, linha=linha, maq=maquinas_da_linha(linha), turnos=turnos))
    return dict(data=data, emissao=emissao, obs=obs, linhas=linhas)


def maquinas_do_aviso(txt):
    """Quais máquinas um aviso do PCP cita (MESPACK F, RETORT, 60L, TOP DOWN, CONDOR...)."""
    T = sem_acento(txt)
    out = []
    for m in re.finditer(r'MESPACK\s+((?:R[12]|[A-F])(?:\s*(?:,|&|E)\s*(?:R[12]|[A-F]))*)', T):
        out += [MP_LETRA[tk] for tk in re.findall(r'\b(R[12]|[A-F])\b', m.group(1))]
    for chave in ('RETORT', '60L', 'SERAC', 'TOP DOWN', 'TOPDOWN', 'FRASCO', 'CONDOR', '1,7KG', 'COPO', '3100', '4KG'):
        if chave in T:
            out += maquinas_da_linha(chave + (' 1 & 2' if chave in ('RETORT', '60L') else ''))
    return list(dict.fromkeys(out))


def tipo_aviso(txt):
    T = sem_acento(txt)
    if re.search(r'MANUTENC|PREVENTIV|CORRETIV|PARADA|QUEBRA|TROCA', T):
        return 'manutencao'
    if re.search(r'CIP|SWAB|TESTE|LIMPEZA|HIGIENI|SANIT|QUALIDADE|ACOMPANHAD', T):
        return 'qualidade'
    return 'info'


def janelas(prog):
    """Para cada máquina: intervalos de produção, horas programadas e janelas sem produção."""
    obs_txt = sem_acento(' '.join(prog['obs']))
    so_cip = set()      # máquinas que, pela observação, são as únicas a fazer CIP no grupo
    for m in re.finditer(r'CIP APENAS NA[S]?\s+((?:MESPACK\s+)?(?:R[12]|[A-F])(?:\s*(?:E|,|&)\s*(?:R[12]|[A-F]))*)', obs_txt):
        for tk in re.findall(r'\b(R[12]|[A-F])\b', m.group(1).replace('MESPACK', '')):
            so_cip.add(MP_LETRA[tk])
    nao_cip = set()     # "NÃO CIPAR MESPACK E": essa máquina fica parada, mas sem fazer CIP
    for m in re.finditer(r'NAO CIPAR\s+((?:MESPACK\s+)?(?:R[12]|[A-F])(?:\s*(?:E|,|&)\s*(?:R[12]|[A-F]))*)', obs_txt):
        for tk in re.findall(r'\b(R[12]|[A-F])\b', m.group(1).replace('MESPACK', '')):
            nao_cip.add(MP_LETRA[tk])
    prog['avisos'] = []
    for o in prog['obs']:
        maq = maquinas_do_aviso(o)
        for cod in re.findall(r'\b111\d{4}\b', o):          # aviso citando o código do produto: máquina da linha dele
            maq += [m for ln in prog['linhas'] if ln['cod'] == cod for m in ln['maq']]
        prog['avisos'].append(dict(txt=o, maq=list(dict.fromkeys(maq)), tipo=tipo_aviso(o)))
    M = {}
    for ln in prog['linhas']:
        for mid in ln['maq']:
            d = M.setdefault(mid, dict(prod=[], ev={}, itens=[]))
            d['itens'].append(dict(desc=ln['desc'], linha=ln['linha'], turnos=ln['turnos']))
            for tr in ln['turnos']:
                if tr['ini'] is not None:
                    d['prod'].append([tr['ini'], tr['fim'], tr['t'], ln['desc']])
                if tr['st'] != 'PROD':
                    st = tr['st']
                    if st == 'CIP' and len(ln['maq']) > 1 and any(x in so_cip for x in ln['maq']) and mid not in so_cip:
                        st = 'PARADA SEM CIP'
                    if st == 'CIP' and mid in nao_cip:
                        st = 'PARADA SEM CIP'
                    d['ev'].setdefault(tr['t'], st)
    for mid, d in M.items():
        d['prod'].sort()
        # horas = união dos intervalos, limitada ao dia (06:30 -> 06:30): nunca passa de 24 h,
        # mesmo com troca de produto no mesmo turno ou um horário lido fora de ordem
        ocupado, fim_ant = 0, 0
        for a, b, *_ in d['prod']:
            a, b = max(0, min(a, 1440)), max(0, min(b, 1440))
            if b <= fim_ant:
                continue
            ocupado += b - max(a, fim_ant)
            fim_ant = b
        d['horas'] = round(ocupado / 60, 2)
        jan = []
        for t, (a, b) in TURNOS.items():
            occ = sorted([(max(a, x), min(b, y)) for x, y, *_ in d['prod'] if x < b and y > a])
            cur = a
            for x, y in occ + [(b, b)]:
                if x - cur >= 15:
                    jan.append(dict(t=t, ini=cur, fim=x, min=x - cur, tipo=d['ev'].get(t, 'SEM PRODUÇÃO')))
                cur = max(cur, y)
            if t in d['ev'] and not any(j['t'] == t for j in jan):
                dur = DUR_EVENTO.get(d['ev'][t])
                if dur:                                   # duração padrão definida: janela no início do turno
                    jan.append(dict(t=t, ini=a, fim=a + dur, min=dur, tipo=d['ev'][t], padrao=True))
                    d['horas'] = round(d['horas'] - dur / 60, 2)
                else:
                    jan.append(dict(t=t, ini=None, fim=None, min=None, tipo=d['ev'][t]))   # evento sem duração
        d['janelas'] = jan
        del d['ev']
    prog['maquinas'] = M
    prog['nao_mapeadas'] = sorted({ln['linha'] for ln in prog['linhas'] if not ln['maq']})
    return prog


if __name__ == '__main__':
  for p in [janelas(x) for x in ler(sys.argv[1])]:
      fmt = lambda m: '—' if m is None else f"{(m + 390) // 60 % 24:02d}:{(m + 390) % 60:02d}"
      print('Programação de', p['data'], '· emitida', p['emissao'], '· obs:', p['obs'])
      for ln in p['linhas']:
          print(f"  {ln['linha']:<20} -> {ln['maq']} | {ln['desc'][:45]:<45} | " +
                ' | '.join(f"T{t['t']} {t['h'] or '—'} {t['st']} {t['qtd']}" for t in ln['turnos']))
      for mid, d in sorted(p['maquinas'].items()):
          print(f"  {mid}: {d['horas']} h programadas; janelas: " +
                '; '.join(f"T{j['t']} {fmt(j['ini'])}–{fmt(j['fim'])} {j['min'] or '?'}min {j['tipo']}" for j in d['janelas']))
      if p['nao_mapeadas']:
          print('  LINHAS NÃO RECONHECIDAS:', p['nao_mapeadas'])
