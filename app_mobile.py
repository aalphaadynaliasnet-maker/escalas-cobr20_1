import streamlit as st
import re, io, warnings, openpyxl, pdfplumber
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from fpdf import FPDF

warnings.filterwarnings("ignore")
st.set_page_config(page_title="Escalas COBR20", page_icon="🚆", layout="wide")

FERIADOS = {
    "2026-01-01","2026-04-03","2026-04-05","2026-04-25","2026-05-01","2026-06-04","2026-06-10","2026-07-04","2026-08-15","2026-10-05","2026-11-01","2026-12-01","2026-12-08","2026-12-25",
    "2027-01-01","2027-03-26","2027-03-28","2027-04-25","2027-05-01","2027-05-27","2027-06-10","2027-07-04","2027-08-15","2027-10-05","2027-11-01","2027-12-01","2027-12-08","2027-12-25",
    "2028-01-01","2028-04-14","2028-04-16","2028-04-25","2028-05-01","2028-06-10","2028-06-15","2028-07-04","2028-08-15","2028-10-05","2028-11-01","2028-12-01","2028-12-08","2028-12-25",
    "2029-01-01","2029-03-30","2029-04-01","2029-04-25","2029-05-01","2029-05-31","2029-06-10","2029-07-04","2029-08-15","2029-10-05","2029-11-01","2029-12-01","2029-12-08","2029-12-25"
}

def processar_linhas_brutas(linhas):
    dados, os_atual, dias_atual = [], None, ""
    col_os, col_dias, col_serv, col_lret, col_pnt, col_repf = 0, 1, 4, 6, 8, 10
    lixo = ("EM VIGOR", "ALTERAÇÃO", "CATEGORIA", "DEPÓSITO", "COIMBRA", "Planeamento", "Versão electrónica", "F.500.02", "Revisão e Venda", "Local", "Hora")
    for r in linhas:
        r = [re.sub(r'\s+', ' ', str(c)).strip() if c is not None else "" for c in r]
        if not any(r): continue
        if any(p in " ".join(r) for p in lixo): continue
        if any("Dias da Semana" in c for c in r):
            for idx, txt in enumerate(r):
                if txt == "OS": col_os = idx
                elif "Dias da Semana" in txt: col_dias = idx
                elif "Designação" in txt: col_serv = idx
                elif "PNT" in txt: col_pnt = idx
            col_lret, col_repf = max(0, col_pnt - 2), col_pnt + 2
            continue
        if len(r) <= col_pnt: continue
        c_os = r.__getitem__(col_os)
        if c_os.replace(".0", "").isdigit():
            val_os = int(float(c_os))
            if 1 <= val_os <= 120: os_atual = val_os
        c_dias = r.__getitem__(col_dias)
        if c_dias and "Dias" not in c_dias: dias_atual = c_dias
        if os_atual:
            serv = r.__getitem__(col_serv) if len(r) > col_serv else ""
            loc_ret = r.__getitem__(col_lret) if len(r) > col_lret else ""
            pnt = r.__getitem__(col_pnt) if len(r) > col_pnt else ""
            rep_f_col = r.__getitem__(col_repf) if len(r) > col_repf else ""
            if c_dias or serv or pnt:
                dados.append({"OS": os_atual, "Dias": c_dias, "Dias_Efet": dias_atual, "Serv": serv, "LocalRet": loc_ret, "PNT": pnt, "RepFora": rep_f_col.lower() == "x"})
    return dados
def ler_ficheiro_carregado(uploaded_file):
    nome, bytes_data = uploaded_file.name.lower(), uploaded_file.read()
    if nome.endswith(".pdf"):
        todas_linhas = []
        with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
            for pagina in pdf.pages:
                for tabela in pagina.extract_tables(): todas_linhas.extend(tabela)
        return processar_linhas_brutas(todas_linhas)
    wb = openpyxl.load_workbook(io.BytesIO(bytes_data), data_only=True)
    for sheet in wb.sheetnames:
        linhas = [["" if c is None else str(c).strip() for c in r] for r in wb[sheet].iter_rows(values_only=True)]
        if any(any("Dias da Semana" in c for c in l) for l in linhas): return processar_linhas_brutas(linhas)
    return []

def contar_repousos_escala(dados):
    return sum(len(set(x["Dias_Efet"] for x in dados if x["OS"] == os_n and x["RepFora"])) for os_n in range(1, 121))

def expandir(regra):
    pares = set()
    for b in str(regra).split():
        m = re.match(r'^([0-9\.]+)([NVFS]*)$', b)
        if not m: continue
        dt, tt = m.groups()
        tipos = list(tt) if tt else ['N','V','F','S']
        dias = range(int(dt.split('..').__getitem__(0)), int(dt.split('..').__getitem__(1)) + 1) if '..' in dt else [int(dt)]
        for d in dias:
            for t in tipos: pares.add((d, t))
    return pares

def calcular_semanas_detalhado(dados, data_ini_str):
    regras_os = {os_n: [] for os_n in range(1, 121)}
    for os_n in range(1, 121):
        por_regra = {}
        for x in [k for k in dados if k["OS"] == os_n]:
            reg = x["Dias_Efet"]
            if not reg: continue
            por_regra.setdefault(reg, 0)
            m_pnt = re.match(r'^(\d{1,2}):(\d{2})', str(x["PNT"]).strip())
            if m_pnt: por_regra[reg] += int(m_pnt.group(1))*60 + int(m_pnt.group(2))
        for reg, mins in por_regra.items(): regras_os[os_n].append((expandir(reg), mins))
    base = datetime.strptime(data_ini_str, "%Y-%m-%d").date()
    lista_semanas = []
    for sem in range(120):
        mins_sem = 0
        for d in range(7):
            dia_abs = sem * 7 + d
            dt = base + timedelta(days=dia_abs)
            de = 1 if dt.weekday() == 6 else dt.weekday() + 2
            tp = 'F' if dt.strftime("%Y-%m-%d") in FERIADOS else ('V' if (dt+timedelta(days=1)).strftime("%Y-%m-%d") in FERIADOS else ('S' if (dt-timedelta(days=1)).strftime("%Y-%m-%d") in FERIADOS else 'N'))
            for pv, mins in regras_os[(dia_abs % 120) + 1]:
                if (de, tp) in pv:
                    mins_sem += mins
                    break
        lista_semanas.append(mins_sem)
    return lista_semanas

def calcular_indicadores_dia(dados):
    regras_os = {os_n: [] for os_n in range(1, 121)}
    for os_n in range(1, 121):
        for x in [k for k in dados if k["OS"] == os_n]:
            reg = x["Dias_Efet"]
            if not reg: continue
            serv_txt = x["Serv"].upper()
            estado = "DESCANSO" if "DESCANSO" in serv_txt else ("SUPRA" if "SUPRA" in serv_txt else "SERVICOS")
            regras_os[os_n].append((expandir(reg), estado))
    contagem = {d: {"SERVICOS": 0, "SUPRA": 0, "DESCANSO": 0} for d in range(1, 8)}
    for d in range(1, 8):
        for os_n in range(1, 121):
            estado_encontrado = "DESCANSO"
            for pv, est in regras_os[os_n]:
                if (d, 'N') in pv:
                    estado_encontrado = est
                    break
            contagem[d][estado_encontrado] += 1
    return contagem

def fmt_hm(mins, com_sinal=False):
    s = "-" if mins < 0 else ("+" if com_sinal and mins > 0 else "")
    m = abs(int(round(mins)))
    return f"{s}{m//60:02d}h{m%60:02d}m"

def comboios(txt):
    limpo = re.sub(r'\(.*?\)', '', str(txt))
    limpo = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', '', limpo)
    return set(re.findall(r'\b\d{3,5}\b', limpo))

def limpar_latin1(txt):
    return str(txt).replace("•", "-").replace("→", "->").encode("latin-1", "replace").decode("latin-1")
def gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, so_os=False):
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    def titulo_sec(txt):
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(31, 78, 120)
        pdf.cell(0, 8, limpar_latin1(txt), new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(20, 20, 20)

    def cabecalho_tabela(cols, largs):
        pdf.set_fill_color(31, 78, 120)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        for c, w in zip(cols, largs):
            pdf.cell(w, 7, limpar_latin1(c), border=1, align="C", fill=True)
        pdf.ln()
        pdf.set_text_color(20, 20, 20)

    titulo_sec("Resumo de OS Alteradas - Escala COBR20" if so_os else "Comparação Completa de Escalas COBR20")
    pdf.ln(2)
    titulo_sec("1. Alterações por Ordem de Serviço (OS)")
    l_os = [18, 45, 107, 107]
    cabecalho_tabela(["OS", "Tipo de Alteração", "Escala Anterior", "Escala Nova (Negrito)"], l_os)

    for os_str, tp, ta, tn, cor_rgb in lista_os_alt:
        linhas_max = max(ta.count("\n") + 1, tn.count("\n") + 1, 1)
        alt_linha = max(7, linhas_max * 4.5)
        if pdf.get_y() + alt_linha > 195:
            pdf.add_page()
            cabecalho_tabela(["OS", "Tipo de Alteração", "Escala Anterior", "Escala Nova (Negrito)"], l_os)
        y_ini = pdf.get_y()
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_os.__getitem__(0), alt_linha, limpar_latin1(os_str), border=1, align="C", fill=True)
        pdf.cell(l_os.__getitem__(1), alt_linha, limpar_latin1(tp), border=1, align="C", fill=True)
        x_ta = pdf.get_x()
        pdf.set_font("Helvetica", "", 7.5)
        pdf.multi_cell(l_os.__getitem__(2), alt_linha / linhas_max, limpar_latin1(ta), border=1, fill=True)
        pdf.set_xy(x_ta + l_os.__getitem__(2), y_ini)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.multi_cell(l_os.__getitem__(3), alt_linha / linhas_max, limpar_latin1(tn), border=1, fill=True)
        pdf.set_y(y_ini + alt_linha)

    if so_os:
        return bytes(pdf.output())

    pdf.add_page()
    titulo_sec("2. Mobilidade de Comboios entre OS")
    l_cb = [25, 45, 50, 50, 53, 54]
    cabecalho_tabela(["Comboio", "Movimento", "Saiu das OS", "Entrou nas OS", "OS Antes", "OS Agora"], l_cb)
    for cb, mv, saiu, entrou, oa, on, cor_rgb in lista_cb_mov:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_cb.__getitem__(0), 6.5, limpar_latin1(cb), border=1, align="C", fill=True)
        pdf.cell(l_cb.__getitem__(1), 6.5, limpar_latin1(mv), border=1, align="C", fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(l_cb.__getitem__(2), 6.5, limpar_latin1(saiu), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_cb.__getitem__(3), 6.5, limpar_latin1(entrou), border=1, fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(l_cb.__getitem__(4), 6.5, limpar_latin1(oa), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_cb.__getitem__(5), 6.5, limpar_latin1(on), border=1, fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("3. Alterações nos Repousos Fora da Sede (REP F)")
    l_rep = [22, 55, 100, 100]
    cabecalho_tabela(["OS", "Situação Repouso", "Repousos Antes", "Repousos Agora"], l_rep)
    for os_s, st_r, ra, rn, cor_rgb in lista_rep_alt:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_rep.__getitem__(0), 6.5, limpar_latin1(os_s), border=1, align="C", fill=True)
        pdf.cell(l_rep.__getitem__(1), 6.5, limpar_latin1(st_r), border=1, align="C", fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(l_rep.__getitem__(2), 6.5, limpar_latin1(ra), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_rep.__getitem__(3), 6.5, limpar_latin1(rn), border=1, fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("4. Balanço Final: Ganhos vs Perdas")
    l_bal = [115, 45, 117]
    cabecalho_tabela(["Indicador", "Antes -> Agora", "Impacto / Saldo"], l_bal)
    for ind, qtd, obs, cor_rgb in balanco_rows:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(l_bal.__getitem__(0), 7, limpar_latin1(ind), border=1, fill=True)
        pdf.cell(l_bal.__getitem__(1), 7, limpar_latin1(qtd), border=1, align="C", fill=True)
        pdf.cell(l_bal.__getitem__(2), 7, limpar_latin1(obs), border=1, fill=True)
        pdf.ln()

    pdf.add_page()
    titulo_sec("5. Distribuição Diária da Escala (Serviços, SUPRA e Descanso)")
    l_dist = [45, 55, 40, 55, 42, 40]
    cabecalho_tabela(["Dia da Semana", "Serviços (Ant -> Nov)", "Dif.", "SUPRA (Ant -> Nov)", "Descanso", "Total OS"], l_dist)
    for d_row in dist_rows:
        pdf.set_fill_color(*d_row.__getitem__(6))
        pdf.set_font("Helvetica", "B", 8.5)
        for idx_c in range(6):
            pdf.cell(l_dist.__getitem__(idx_c), 7, limpar_latin1(d_row.__getitem__(idx_c)), border=1, align=("L" if idx_c==0 else "C"), fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("6. Cômputo a cada 12 Semanas (Rotação 120 Semanas)")
    l_sem = [35, 55, 62, 62, 63]
    cabecalho_tabela(["Bloco", "Semanas", "Horas Anterior", "Horas Nova", "Diferença (12 Sem.)"], l_sem)
    for s_row in sem_rows:
        pdf.set_fill_color(255, 235, 156)
        pdf.set_font("Helvetica", "B", 8.5)
        for idx_c in range(5):
            pdf.cell(l_sem.__getitem__(idx_c), 7, limpar_latin1(s_row.__getitem__(idx_c)), border=1, align="C", fill=True)
        pdf.ln()

    return bytes(pdf.output())
def processar_comparacao_mobile(d_ant, d_nov, data_ini_str):
    sems_ant, sems_nov = calcular_semanas_detalhado(d_ant, data_ini_str), calcular_semanas_detalhado(d_nov, data_ini_str)
    c_ant, c_nov, os_alt, os_gan, os_sup, rep_alt = {}, {}, 0, 0, 0, []
    lista_os_alt, lista_cb_mov, lista_rep_alt = [], [], []

    for os_n in range(1, 121):
        sa, sn = [x for x in d_ant if x["OS"] == os_n], [x for x in d_nov if x["OS"] == os_n]
        ta = "\n".join(f"- [{x['Dias']}] {x['Serv']} ({x['PNT'][:5]}){' [REP FORA: '+x['LocalRet']+']' if x['RepFora'] else ''}" for x in sa if x["Dias"] or x["Serv"])
        tn = "\n".join(f"- [{x['Dias']}] {x['Serv']} ({x['PNT'][:5]}){' [REP FORA: '+x['LocalRet']+']' if x['RepFora'] else ''}" for x in sn if x["Dias"] or x["Serv"])
        rep_a = sorted(set(f"{x['Dias_Efet']} ({x['LocalRet']})" for x in sa if x["RepFora"]))
        rep_n = sorted(set(f"{x['Dias_Efet']} ({x['LocalRet']})" for x in sn if x["RepFora"]))
        if rep_a != rep_n: rep_alt.append((os_n, rep_a, rep_n))
        for x in sa:
            for cb in comboios(x["Serv"]): c_ant.setdefault(cb, set()).add(os_n)
        for x in sn:
            for cb in comboios(x["Serv"]): c_nov.setdefault(cb, set()).add(os_n)
        if ta != tn:
            os_alt += 1
            if "SUPRA" not in ta and "SUPRA" in tn: tp, cor, os_sup = "PASSOU A SUPRA", (255, 235, 156), os_sup + 1
            elif "SUPRA" in ta and "SUPRA" not in tn: tp, cor, os_gan = "GANHOU SERVIÇO", (198, 239, 206), os_gan + 1
            else: tp, cor = "AJUSTE HORÁRIO / COMBOIO", (217, 225, 242)
            lista_os_alt.append((f"OS {os_n:03d}", tp, ta, tn, cor))

    cb_nov, cb_sup, cb_trf = 0, 0, 0
    for cb in sorted(set(c_ant) | set(c_nov), key=int):
        oa, on = sorted(c_ant.get(cb, set())), sorted(c_nov.get(cb, set()))
        if oa != on:
            saiu = ", ".join(f"OS {x}" for x in sorted(set(oa)-set(on))) or "-"
            entrou = ", ".join(f"OS {x}" for x in sorted(set(on)-set(oa))) or "-"
            if not oa: mv, cor, cb_nov = "NOVO NA ESCALA", (198, 239, 206), cb_nov + 1
            elif not on: mv, cor, cb_sup = "SUPRIMIDO", (255, 199, 206), cb_sup + 1
            else: mv, cor, cb_trf = "TRANSFERIDO", (255, 235, 156), cb_trf + 1
            lista_cb_mov.append((cb, mv, saiu, entrou, ", ".join(f"OS {x}" for x in oa) or "-", ", ".join(f"OS {x}" for x in on) or "-", cor))

    for os_n, ra, rn in rep_alt:
        if len(rn) > len(ra): st_r, cor = "GANHOU REPOUSO", (198, 239, 206)
        elif len(rn) < len(ra): st_r, cor = "PERDEU REPOUSO", (255, 199, 206)
        else: st_r, cor = "ALTEROU REP FORA", (217, 225, 242)
        lista_rep_alt.append((f"OS {os_n:03d}", st_r, ", ".join(ra) if ra else "Sem rep. fora", ", ".join(rn) if rn else "Sem rep. fora", cor))

    m_ant, m_nov = sum(sems_ant) / 120.0, sum(sems_nov) / 120.0
    rep_tot_ant, rep_tot_nov = contar_repousos_escala(d_ant), contar_repousos_escala(d_nov)
    dif_rep, dm, scb, sos = rep_tot_nov - rep_tot_ant, m_nov - m_ant, len(c_nov) - len(c_ant), os_gan - os_sup
    txt_rep = f"GANHOU {dif_rep} repouso(s) (+{dif_rep})" if dif_rep > 0 else (f"PERDEU {abs(dif_rep)} repouso(s) ({dif_rep})" if dif_rep < 0 else "Manteve repousos (0)")
    cor_rep = (198, 239, 206) if dif_rep > 0 else ((255, 199, 206) if dif_rep < 0 else (217, 225, 242))

    balanco_rows = [
        ("Total de Repousos Fora da Sede (REP F)", f"{rep_tot_ant} -> {rep_tot_nov}", txt_rep, cor_rep),
        ("Carga Horária Média Semanal", f"{fmt_hm(m_ant)} -> {fmt_hm(m_nov)}", f"Diferença semanal: {fmt_hm(dm, True)}", (217, 225, 242)),
        ("Comboios distintos na Escala", f"{len(c_ant)} -> {len(c_nov)}", f"Novos: +{cb_nov} | Suprimidos: -{cb_sup} | Transf.: {cb_trf}", (245, 245, 245)),
        ("SALDO LÍQUIDO DE COMBOIOS", f"{scb:+d}", "Ganho/Perda líquida de circulações", (198, 239, 206) if scb>0 else ((255, 199, 206) if scb<0 else (217, 225, 242))),
        ("Total de OS alteradas", f"{os_alt} / 120", f"Ganharam serviço: +{os_gan} | Passaram a SUPRA: -{os_sup}", (217, 225, 242)),
        ("SALDO LÍQUIDO DE OS COM SERVIÇO", f"{sos:+d}", "Balanço entre OS ganhas e SUPRA", (198, 239, 206) if sos>0 else ((255, 199, 206) if sos<0 else (217, 225, 242))),
    ]

    ind_ant, ind_nov = calcular_indicadores_dia(d_ant), calcular_indicadores_dia(d_nov)
    nomes_dias = ("Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado")
    dist_rows = []
    for d_idx in range(1, 8):
        s_a, s_n = ind_ant[d_idx]["SERVICOS"], ind_nov[d_idx]["SERVICOS"]
        u_a, u_n = ind_ant[d_idx]["SUPRA"], ind_nov[d_idx]["SUPRA"]
        desc_n, dif_s = ind_nov[d_idx]["DESCANSO"], s_n - s_a
        cor_l = (198, 239, 206) if dif_s > 0 else ((255, 199, 206) if dif_s < 0 else (248, 248, 248))
        dist_rows.append((nomes_dias.__getitem__(d_idx-1), f"{s_a} -> {s_n}", f"{dif_s:+d}", f"{u_a} -> {u_n}", str(desc_n), str(s_n+u_n+desc_n), cor_l))

    sem_rows = []
    for b_idx in range(10):
        ini_s, fim_s = b_idx * 12, (b_idx + 1) * 12
        sa12, sn12 = sum(sems_ant[ini_s:fim_s]), sum(sems_nov[ini_s:fim_s])
        sem_rows.append((f"BLOCO {b_idx+1}", f"Sem {ini_s+1} a {fim_s}", fmt_hm(sa12), fmt_hm(sn12), fmt_hm(sn12 - sa12, True)))

    pdf_completo = gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, so_os=False)
    pdf_so_os = gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, so_os=True)
    resumo = {"rep_ant": rep_tot_ant, "rep_nov": rep_tot_nov, "dif_rep": dif_rep, "scb": scb, "sos": sos, "h_ant": fmt_hm(m_ant), "h_nov": fmt_hm(m_nov)}
    return pdf_completo, pdf_so_os, resumo

st.title("🚆 Escalas COBR20 — Mobile")
st.caption("Comparador oficial de alterações (Excel ou PDF) otimizado para telemóvel.")

arq_ant = st.file_uploader("1. Escala Anterior (ex: 61ª Alteração)", type=["xlsx", "pdf"])
arq_nov = st.file_uploader("2. Escala Nova (ex: 62ª Alteração)", type=["xlsx", "pdf"])
data_ini = st.date_input("Início da Escala (Dom)", value=datetime(2026, 11, 22))

if arq_ant and arq_nov:
    if st.button("🔍 Comparar Escalas", type="primary", use_container_width=True):
        with st.spinner("A processar escalas..."):
            d_a, d_n = ler_ficheiro_carregado(arq_ant), ler_ficheiro_carregado(arq_nov)
            p_comp, p_os, res = processar_comparacao_mobile(d_a, d_n, data_ini.strftime("%Y-%m-%d"))
            st.session_state["mob_res"] = (p_comp, p_os, res)

if "mob_res" in st.session_state:
    p_comp, p_os, res = st.session_state["mob_res"]
    st.success("Comparação concluída!")
    c1, c2 = st.columns(2)
    c1.metric("Repousos Fora", f"{res['rep_ant']} → {res['rep_nov']}", f"{res['dif_rep']:+d}")
    c2.metric("Saldo Comboios", f"{res['scb']:+d}")
    c3, c4 = st.columns(2)
    c3.metric("Saldo OS Serviço", f"{res['sos']:+d}")
    c4.metric("Média Horas/Sem", res["h_nov"], f"Antes: {res['h_ant']}")

    st.download_button("📕 Descarregar PDF Completo", data=p_comp, file_name="Comparacao_Escalas_COBR20.pdf", mime="application/pdf", use_container_width=True)
    st.download_button("📋 Descarregar PDF Só OS Alteradas", data=p_os, file_name="OS_Alteradas_COBR20.pdf", mime="application/pdf", use_container_width=True)
