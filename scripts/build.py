"""Gera o painel de Preventiva das Envasadoras (index.html) a partir de:
  - uploads/*.xlsx  -> abas Base, Preventiva e Execuções (planilha de tagueamento)
  - BASE_PCM        -> falhas corretivas (planilha do MTBF/MTTR, baixada pelo Actions)
  - programacao/*.pdf -> Programação de Produção Diária do PCP

Uso: python scripts/build.py --pcm caminho/PCM.xlsm [--out index.html]
"""
import argparse, glob, json, math, os, re, sys, unicodedata
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
import parse_prog  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------- configuração
FAMILIAS_ENVASE = ['MSP', 'MPS', 'ENC', 'ENA']                       # famílias que entram no painel
AREAS = {'LINHA ATOMATADOS': 'Atomatados', 'LINHA POLPA': 'Polpa', 'LINHA VEGETAIS': 'Vegetais'}
# nome do equipamento na BASE_PCM -> id da máquina no painel
PCM_MAQ = {'MESPACK A': 'ATO-MP01', 'MESPACK B': 'ATO-MP02', 'MESPACK C': 'ATO-MP03', 'MESPACK D': 'ATO-MP04',
           'MESPACK E': 'ATO-MP05', 'MESPACK F': 'ATO-MP07', 'MESPACK R1': 'VEG-MP06', 'MESPACK R2': 'VEG-MP08',
           'ENCHEDEIRA TOP DOWN': 'ATO-EC08', 'ENCHEDEIRA COPO/LATA': 'ATO-EC06', '60L1': 'VEG-EC02', '60L2': 'VEG-EC03'}
IGNORAR_TAG = {'ETH'}                                                # TAGs que não são envasadoras
FREQ_DIAS = {'DIÁRIA': 1, 'DIARIA': 1, 'SEMANAL': 7, 'QUINZENAL': 15, 'MENSAL': 30, 'BIMESTRAL': 60,
             'TRIMESTRAL': 90, 'SEMESTRAL': 180, 'ANUAL': 365}


def norm(s):
    s = ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip().upper()


def col(df, *nomes):
    """Acha a coluna pelo nome, ignorando acento, caixa e espaços."""
    alvo = {norm(n) for n in nomes}
    for c in df.columns:
        if norm(c) in alvo:
            return c
    return None


def txt(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def dt(v):
    if v is None or v == '' or (isinstance(v, float) and math.isnan(v)):
        return None
    try:
        d = pd.to_datetime(v, dayfirst=True, errors='coerce')
        return None if pd.isna(d) else d.date()
    except Exception:
        return None


def num(v):
    try:
        f = float(str(v).replace(',', '.'))
        return None if math.isnan(f) else f
    except Exception:
        return None


# ---------------------------------------------------------------- planilha de tagueamento
def ler_base(xlsx):
    df = pd.read_excel(xlsx, sheet_name='Base', header=7, usecols=range(14)).dropna(subset=['TAG'])
    df['TAG'] = df['TAG'].astype(str).str.strip()
    df['rg'] = df['RG COMPONENTE'].map(txt)
    return df


def maquinas(df):
    env = df[df.FAMILIA.isin(FAMILIAS_ENVASE) & (df.STATUS == 'ATIVO') & df['ÁREA'].isin(AREAS)].copy()
    env['A'] = env['ÁREA'].map(AREAS)
    ms, comps = [], []
    for (a, t), g in env.groupby(['A', 'TAG']):
        if t in IGNORAR_TAG:
            continue
        mid = f'{a[:3].upper()}-{t}'
        ms.append(dict(id=mid, tag=t, nome=g['EQUIPAMENTO'].astype(str).str.strip().mode()[0], area=a,
                       linha=str(g.LINHA.mode()[0]).strip() if g.LINHA.notna().any() else '',
                       crit='A' if (g.CRITICIDADE == 'A').any() else 'B', ncomp=len(g)))
        comps += [(mid, r) for r in g.rg]
    return ms, comps


def deptos(txt_):
    """'Mecânica / Elétrica', 'Elétrica/Manutenção', 'MEC+ELE'... -> ['Mecânica', 'Elétrica'] (ordem fixa)."""
    t = norm(txt_)
    out = []
    if 'MEC' in t or 'MANUT' in t:
        out.append('Mecânica')
    if 'ELE' in t:
        out.append('Elétrica')
    if 'AUT' in t:
        out.append('Automação')
    return out or ['Mecânica']


def ler_planos(xlsx, base, ms):
    p = pd.read_excel(xlsx, sheet_name='Preventiva')
    c = {k: col(p, *v) for k, v in dict(
        tag=['TAG'], rg=['RG'], comp=['EQUIPAMENTO'], at=['ATIVIDADE'], tipo=['TIPO MANUTENÇÃO'], freq=['FREQ.', 'FREQ'],
        per=['PERIODO EM DIAS'], ult=['ULTIMA EXECUÇÃO'], os=['OS'], disc=['DEPARTAMENTO', 'DISCIPLINA'],
        h=['HORAS ESTIMADAS'], cond=['CONDIÇÃO'], mat=['MATERIAIS'], base=['BASE']).items()}
    rg_maq = {}                                     # RG -> id da máquina, pela aba Base
    ids = {m['id'] for m in ms}
    for _, r in base.iterrows():
        a = AREAS.get(r['ÁREA'])
        if a and f"{a[:3].upper()}-{r.TAG}" in ids:
            rg_maq.setdefault(r.rg, f"{a[:3].upper()}-{r.TAG}")
    por_tag = {}
    for m in ms:
        por_tag.setdefault(m['tag'], []).append(m['id'])
    planos, fora = [], []
    for _, r in p.iterrows():
        at = txt(r[c['at']]) if c['at'] else ''
        if not at:
            continue
        tag, rg = txt(r[c['tag']]), txt(r[c['rg']])
        mid = rg_maq.get(rg) or (por_tag[tag][0] if len(por_tag.get(tag, [])) == 1 else None)
        chave = f'{tag} | {rg} | {at}'
        if not mid:
            fora.append(chave)
            continue
        freq = txt(r[c['freq']]) if c['freq'] else ''
        per = num(r[c['per']]) if c['per'] else None
        per = int(per) if per else FREQ_DIAS.get(norm(freq))
        if not per:
            fora.append(chave + ' (sem frequência)')
            continue
        disc = txt(r[c['disc']]) if c['disc'] else ''
        cond = norm(txt(r[c['cond']])) if c['cond'] else ''
        planos.append(dict(id=chave, m=mid, rg=rg, comp=txt(r[c['comp']]) if c['comp'] else '', at=at,
                           tipo=txt(r[c['tipo']]) if c['tipo'] else '', freq=freq or f'{per} dias', per=per,
                           discs=deptos(disc), disc=' + '.join(deptos(disc)),
                           h=(num(r[c['h']]) if c['h'] else None) or 0, cond='Parada' if 'PARADA' in cond else 'Rodando',
                           mat=txt(r[c['mat']]) if c['mat'] else '', base=txt(r[c['base']]) if c['base'] else 'Dias',
                           ult0=dt(r[c['ult']]) if c['ult'] else None, os=txt(r[c['os']]) if c['os'] else '', real=False))
    return planos, fora


def ler_execucoes(xlsx, planos):
    try:
        e = pd.read_excel(xlsx, sheet_name='Execuções')
    except ValueError:
        return []
    c = {k: col(e, *v) for k, v in dict(plano=['PLANO'], prog=['DATA PROGRAMADA'], ex=['DATA EXECUTADA'],
                                         hr=['HORAS REAIS'], os=['OS'], tu=['TURNO']).items()}
    ids = {p['id'] for p in planos}
    out = []
    for _, r in e.iterrows():
        pid = txt(r[c['plano']])
        prog = dt(r[c['prog']]) or dt(r[c['ex']])
        if pid not in ids or not prog:
            continue
        ex = dt(r[c['ex']])
        out.append([pid, prog.isoformat(), ex.isoformat() if ex else '', (num(r[c['hr']]) if c['hr'] else None) or 0,
                    txt(r[c['os']]) if c['os'] else '', txt(r[c['tu']]) if c['tu'] else ''])
    # última execução: maior data executada do histórico; senão, a coluna ULTIMA EXECUÇÃO
    for p in planos:
        exs = [x[2] for x in out if x[0] == p['id'] and x[2]]
        last = max(exs) if exs else (p['ult0'].isoformat() if p['ult0'] else '')
        p['ult'] = last
        oss = [x[4] for x in out if x[0] == p['id'] and x[2] and x[4]]
        if oss:
            p['os'] = oss[-1]
        del p['ult0']
    return out


# ---------------------------------------------------------------- falhas (BASE_PCM)
def ler_pcm(pcm_path, base):
    b = pd.read_excel(pcm_path, sheet_name='BASE_PCM', usecols=range(19)).dropna(subset=['Data Inicio'])
    b['Data Inicio'] = pd.to_datetime(b['Data Inicio'], errors='coerce')
    b = b.dropna(subset=['Data Inicio'])
    b['rg'] = b.TAG.map(txt)
    b['rgl'] = b.rg.str.lower()
    bl = base.assign(rgl=base.rg.str.lower()).drop_duplicates('rgl').set_index('rgl')
    b['tagb'] = b.rgl.map(bl.TAG)
    b['mid'] = b.Equipamento.astype(str).str.strip().str.upper().map(PCM_MAQ)
    b['h'] = pd.to_numeric(b['Tempo Parada (h decimal)'], errors='coerce').fillna(0)
    b['f'] = pd.to_numeric(b['Falhas'], errors='coerce').fillna(1)
    b['MOTIVO'] = b.MOTIVO.astype(str).str.strip()
    return b


def agrega(s):
    return dict(f=int(s.f.sum()), h=round(float(s.h.sum()), 2),
                dep={k: round(float(v), 2) for k, v in s.groupby('Departamento').h.sum().items()},
                depf={k: int(v) for k, v in s.groupby('Departamento').f.sum().items()},
                mot={k: [[a, int(n)] for a, n in g.groupby('MOTIVO').f.sum().sort_values(ascending=False).head(3).items()]
                     for k, g in s.groupby('Departamento')})


def falhas(b):
    b = b.copy()
    b['mes'] = b['Data Inicio'].dt.strftime('%Y-%m')
    dias = b.groupby('mes')['Data Inicio'].nunique()
    meses = [k for k, v in dias.items() if v >= 10][-6:]
    fal = {mid: {m: agrega(g[g.mes == m]) for m in meses} for mid, g in b[b.mid.notna()].groupby('mid')}
    last = b['Data Inicio'].max().normalize()
    sem = []
    for i in range(7, -1, -1):
        fim = last - pd.Timedelta(days=7 * i)
        sem.append(dict(k=f'w{7 - i}', ini=str((fim - pd.Timedelta(days=6)).date()), fim=str(fim.date())))
    b['wk'] = None
    for w in sem:
        b.loc[(b['Data Inicio'] >= w['ini']) & (b['Data Inicio'] <= w['fim']), 'wk'] = w['k']
    falw = {mid: {w['k']: agrega(g[g.wk == w['k']]) for w in sem} for mid, g in b[b.mid.notna() & b.wk.notna()].groupby('mid')}
    c90 = b[(b['Data Inicio'] > last - pd.Timedelta(days=90)) & b.mid.notna()]
    comp = []
    for (mid, rg), g in c90.groupby(['mid', 'rg']):
        desc = g.Componente.dropna().astype(str).str.strip().mode()
        comp.append(dict(m=mid, rg=rg, tagb=txt(g.tagb.iloc[0]), d=desc[0] if len(desc) else rg, f=int(g.f.sum()),
                         h=round(float(g.h.sum()), 2), mot=g.MOTIVO.value_counts().head(2).index.tolist(),
                         dep=g.Departamento.mode()[0] if g.Departamento.notna().any() else ''))
    comp.sort(key=lambda x: -x['h'])
    return dict(fal=fal, meses=meses, diasMes={k: int(dias[k]) for k in meses}, falw=falw, semanas=sem,
                diasSem={w['k']: int(b[b.wk == w['k']]['Data Inicio'].nunique()) or 1 for w in sem},
                comp=comp[:150], pcmPeriodo=[str(b['Data Inicio'].min().date()), str(last.date())], pcmReg=int(len(b)))


# ---------------------------------------------------------------- qualidade da base
def qualidade(xlsx, df, b):
    env = df[df.FAMILIA.isin(FAMILIAS_ENVASE) & (df.STATUS == 'ATIVO') & df['ÁREA'].isin(AREAS)].copy()
    env['A'] = env['ÁREA'].map(AREAS)
    R = []
    for t, g in env.groupby('TAG'):
        if g.A.nunique() > 1:
            R.append(('TAG repetida em áreas diferentes', t, ' / '.join(
                f"{a}: {', '.join(sorted(set(x.EQUIPAMENTO.astype(str).str.strip())))}" for a, x in g.groupby('A'))))
    for (a, t), g in env.groupby(['A', 'TAG']):
        if t in IGNORAR_TAG:
            continue
        s = sorted(set(g.EQUIPAMENTO.astype(str).str.strip()))
        if len(s) > 1:
            R.append(('Mesmo TAG com nomes diferentes', t, f'{a}: ' + ' · '.join(s)))
        if g.CRITICIDADE.nunique() > 1:
            R.append(('Criticidade A e B no mesmo equipamento', t,
                      f'{s[0]} ({a}): ' + ', '.join(f'{k}={v}' for k, v in g.CRITICIDADE.value_counts().items())))
        if len(g) <= 4:
            R.append(('Poucos componentes cadastrados', t, f'{s[0]} ({a}): {len(g)} componente(s)'))
    d = df[df.rg.duplicated(keep=False) & (df.rg != '')]
    for rg, g in d.groupby('rg'):
        R.append(('RG repetido na base', rg, ' / '.join(f"{r.TAG}: {txt(r['DESCRICAO COMPONENTE'])[:55]}" for _, r in g.iterrows())))
    ws = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)['Base']
    areas = {}
    for r in ws.iter_rows(min_row=9, max_row=8 + len(df) + 50, max_col=11, values_only=True):
        if r[5] and isinstance(r[10], str) and r[10].startswith('#'):
            areas.setdefault(str(r[2]), []).append(str(r[5]))
    for a, ts in areas.items():
        R.append(('Centro de custo #N/A (área fora da aba Centro de Custo)', a, ', '.join(ts)))
    if ws.max_row and ws.max_row > len(df) + 5000:
        R.append(('Planilha inflada', 'Base', f'A aba Base está formatada até a linha {ws.max_row:,}'.replace(',', '.')
                  + f', mas só {len(df):,} linhas têm TAG. Apagar as linhas vazias abaixo da tabela reduz o arquivo'.replace(',', '.')
                  + ' e deixa o envio ao GitHub mais rápido.'))
    if b is not None:
        nm = b[b.tagb.isna()]
        vc = nm.rg.value_counts()
        if len(nm):
            R.append(('Código da BASE_PCM sem cadastro na base de TAGs', '—',
                      f'{len(nm)} de {len(b)} registros. ' + ', '.join(f'{k} ({v})' for k, v in vc.head(12).items())))
        mf = b[(b.mid == 'ATO-MP07') & (b.tagb == 'MP05')]
        if len(mf):
            R.append(('Falhas lançadas em componente de outra máquina', 'MP05',
                      f'{len(mf)} paradas da MESPACK F usam RGs cadastrados como MP05 (Mespack E). Ex.: '
                      + ', '.join(mf.rg.value_counts().head(5).index)))
    return [dict(tipo=a, tag=t, det=c) for a, t, c in R]


# ---------------------------------------------------------------- programação (PDFs)
def programacoes():
    progs = {}
    for f in sorted(glob.glob(os.path.join(ROOT, 'programacao', '*.pdf')), key=os.path.getmtime):
        try:
            for p in parse_prog.ler(f):
                p = parse_prog.janelas(p)
                p['arquivo'] = os.path.basename(f)
                progs[p['data']] = p                    # o arquivo mais recente vence (reenvio/correção)
        except Exception as e:                          # PDF fora do padrão não derruba o painel
            print('AVISO: não consegui ler', f, '-', e)
    os.makedirs(os.path.join(ROOT, 'data', 'programacao'), exist_ok=True)
    for d, p in progs.items():
        json.dump(p, open(os.path.join(ROOT, 'data', 'programacao', f'{d}.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    ultimas = sorted(progs)[-7:]
    return {d: progs[d] for d in ultimas}


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pcm', required=True)
    ap.add_argument('--xlsx', default=None)
    ap.add_argument('--out', default=os.path.join(ROOT, 'index.html'))
    a = ap.parse_args()
    cands = glob.glob(os.path.join(ROOT, 'uploads', '*.xls*'))
    if not a.xlsx and not cands:
        print('Nenhuma planilha em uploads/ ainda. Nada a gerar.')
        return
    xlsx = a.xlsx or max(cands, key=os.path.getmtime)
    print('Planilha:', xlsx)
    base = ler_base(xlsx)
    ms, comps = maquinas(base)
    planos, fora = ler_planos(xlsx, base, ms)
    execs = ler_execucoes(xlsx, planos)
    b = ler_pcm(a.pcm, base) if a.pcm and os.path.exists(a.pcm) else None
    F = falhas(b) if b is not None else dict(fal={}, meses=[], diasMes={}, falw={}, semanas=[], diasSem={}, comp=[],
                                              pcmPeriodo=['', ''], pcmReg=0)
    progs = programacoes()
    issues = qualidade(xlsx, base, b)
    if fora:
        issues.append(dict(tipo='Planos que não entram neste painel', tag='—',
                           det=f'{len(fora)} plano(s) da aba Preventiva são de outros equipamentos ou estão sem frequência: '
                               + '; '.join(fora[:8])))
    D = dict(machines=ms, comps=len(comps), plans=planos, execs=execs, issues=issues, progs=progs,
             prog=progs[max(progs)] if progs else None,
             gerado=datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M'), **F)
    tpl = open(os.path.join(ROOT, 'scripts', 'template.html'), encoding='utf-8').read()
    html = tpl.replace('/*DATA*/', json.dumps(D, ensure_ascii=False, separators=(',', ':'), default=str))
    open(a.out, 'w', encoding='utf-8').write(html)
    print(f'OK: {len(ms)} máquinas, {len(planos)} planos, {len(execs)} execuções, {F["pcmReg"]} registros BASE_PCM, '
          f'{len(progs)} programação(ões), {len(issues)} pendências -> {a.out}')


if __name__ == '__main__':
    main()
