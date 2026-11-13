from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import PatternFill, Font
from openai import OpenAI
import os
import time
import re

# Carrega as variáveis do arquivo .env
load_dotenv()

st.set_page_config(page_title="Agente IA Cirúrgico - Conciliação Contábil", layout="wide")
st.header("🤖 Agente Auditor IA: Conciliação Emparelhada e Reorganizada por Título")

# --- UPLOAD DE ARQUIVOS NA INTERFACE PRINCIPAL ---
st.subheader("📁 Envio de Arquivos")
col_up1, col_up2, col_up3 = st.columns(3)
with col_up1:
    arquivo_razao = st.file_uploader("1. Arquivo do Razão (Obrigatório)", type=["xlsx", "xls"], key="up_razao")
with col_up2:
    arquivo_apb = st.file_uploader("2. Planilha APB (Opcional)", type=["xlsx", "xls"], key="up_apb")
with col_up3:
    arquivo_pendencias = st.file_uploader("3. Pendências Mês Anterior (Opcional)", type=["xlsx", "xls"], key="up_pend")

if arquivo_razao is not None:
    df_razao = pd.read_excel(arquivo_razao)
    df_apb = pd.read_excel(arquivo_apb) if arquivo_apb is not None else None
    df_pendencias = pd.read_excel(arquivo_pendencias) if arquivo_pendencias is not None else None

    st.success("✅ Arquivo principal do Razão carregado com sucesso!")

    # --- MAPEAMENTO DE COLUNAS ---
    st.markdown("---")
    st.subheader("⚙️ Configuração e Mapeamento de Colunas")

    rc1, rc2, rc3, rc4 = st.columns(4)
    with rc1:
        col_debito_raz = st.selectbox("Coluna de Débito:", options=list(df_razao.columns), key="r_deb")
    with rc2:
        col_credito_raz = st.selectbox("Coluna de Crédito:", options=list(df_razao.columns), key="r_cred")
    with rc3:
        col_titulo_raz = st.selectbox("Coluna de Título / Histórico:", options=list(df_razao.columns), key="r_tit")
    with rc4:
        col_conta_raz = st.selectbox("Conta Contábil (Opcional):", options=["Nenhuma"] + list(df_razao.columns), key="r_conta")

    if col_conta_raz != "Nenhuma":
        contas_raz = df_razao[col_conta_raz].dropna().unique().tolist()
        filtro_conta_raz = st.selectbox("Filtrar Conta Específica no Razão:", options=["Todas"] + contas_raz, key="f_r_conta")
        if filtro_conta_raz != "Todas":
            df_razao = df_razao[df_razao[col_conta_raz] == filtro_conta_raz]

    # --- PROCESSAMENTO COM REORGANIZAÇÃO DE PARES E DUPLA PASSADA DE REVISÃO ---
    if st.button("🚀 Reorganizar Valores, Emparelhar e Executar Agente IA", key="btn_processar_ia_reorganizado"):
        openai_key = os.getenv("OPENAI_API_KEY")
        base_url = os.getenv("BASE_URL")
        modelo = os.getenv("MODELO_OPENAI", "gpt-4o-mini")

        if not openai_key:
            st.error("❌ A chave `OPENAI_API_KEY` não foi encontrada no arquivo `.env`.")
        else:
            with st.spinner("🤖 Executando primeira passada e segunda mão de revisão nas sobras..."):
                try:
                    client_args = {"api_key": openai_key}
                    if base_url:
                        client_args["base_url"] = base_url
                    client = OpenAI(**client_args)

                    df_r = df_razao.copy()

                    # Tratamento numérico seguro preservando floats nativos do Excel
                    def parse_val(v):
                        if pd.isna(v):
                            return 0.0
                        if isinstance(v, (int, float)):
                            return abs(float(v))
                        s = str(v).strip()
                        if s in ["-", "", " - ", "nan", "None"]:
                            return 0.0
                        s = re.sub(r'[R$\s]', '', s)
                        if "," in s and "." in s:
                            if s.rfind(",") > s.rfind("."):
                                s = s.replace(".", "").replace(",", ".")
                            else:
                                s = s.replace(",", "")
                        elif "," in s:
                            s = s.replace(",", ".")
                        try:
                            return abs(float(s))
                        except:
                            return 0.0

                    df_r["_val_deb_"] = df_r[col_debito_raz].apply(parse_val)
                    df_r["_val_cred_"] = df_r[col_credito_raz].apply(parse_val)

                    # Extração rigorosa de chave de documento
                    def extrair_doc(texto):
                        if pd.isna(texto):
                            return "GERAL"
                        matches = re.findall(r'(?:DOC|NR|TIT|FAT|NF|DP|RC|BORD)[:/\s\-]*([0-9A-Z\-\/]+)', str(texto), re.IGNORECASE)
                        if matches:
                            doc_limpo = matches[-1].strip().lstrip("0")
                            return doc_limpo if doc_limpo else "0"
                        numeros = re.findall(r'\b[0-9]{3,12}\b', str(texto))
                        if numeros:
                            doc_limpo = numeros[0].lstrip("0")
                            return doc_limpo if doc_limpo else "0"
                        return str(texto).strip()[:25]

                    df_r["_doc_extraido_"] = df_r[col_titulo_raz].apply(extrair_doc)

                    # ETAPA 1: PASSADA PRINCIPAL (Agrupamento por Documento)
                    blocos_conciliados = []
                    blocos_pendentes = []

                    for doc, grupo in df_r.groupby("_doc_extraido_", sort=False):
                        grupo_ordenado = grupo.sort_values(by="_val_cred_", ascending=False).copy()
                        tot_cred = round(grupo_ordenado["_val_cred_"].sum(), 2)
                        tot_deb = round(grupo_ordenado["_val_deb_"].sum(), 2)

                        if doc != "GERAL" and len(grupo_ordenado) > 1 and abs(tot_cred - tot_deb) <= 0.01 and tot_cred > 0:
                            grupo_ordenado["_pre_conciliado_"] = True
                            blocos_conciliados.append(grupo_ordenado)
                        else:
                            grupo_ordenado["_pre_conciliado_"] = False
                            blocos_pendentes.append(grupo_ordenado)

                    # ETAPA 2: SEGUNDA MÃO DE REVISÃO NAS SOBRAS (Varredura de valores iguais cruzados)
                    if blocos_pendentes:
                        df_pend_all = pd.concat(blocos_pendentes, ignore_index=True)
                        novos_conciliados_revisao = []
                        ainda_pendentes = []

                        cred_pend = df_pend_all[df_pend_all["_val_cred_"] > 0]
                        deb_pend = df_pend_all[df_pend_all["_val_deb_"] > 0]

                        usados_revisao = set()
                        for idx_c, row_c in cred_pend.iterrows():
                            if idx_c in usados_revisao:
                                continue
                            val_alvo = round(row_c["_val_cred_"], 2)
                            
                            # Filtro corrigido com parênteses individuais para evitar conflito bitwise_and
                            matches_deb = deb_pend[
                                ((deb_pend["_val_deb_"] - val_alvo).abs() <= 0.01) & 
                                (~deb_pend.index.isin(usados_revisao))
                            ]
                            
                            if not matches_deb.empty:
                                idx_d = matches_deb.index[0]
                                usados_revisao.add(idx_c)
                                usados_revisao.add(idx_d)
                                par_rev = df_pend_all.loc[[idx_c, idx_d]].copy()
                                par_rev = par_rev.sort_values(by="_val_cred_", ascending=False)
                                par_rev["_pre_conciliado_"] = True
                                novos_conciliados_revisao.append(par_rev)

                        if novos_conciliados_revisao:
                            blocos_conciliados.extend(novos_conciliados_revisao)
                            ainda_pendentes = [df_pend_all[~df_pend_all.index.isin(usados_revisao)]] if len(usados_revisao) < len(df_pend_all) else []
                        else:
                            ainda_pendentes = [df_pend_all]

                        blocos_pendentes = ainda_pendentes

                    # Monta o DataFrame final organizado
                    dfs_ordenados = []
                    if blocos_conciliados:
                        dfs_ordenados.extend(blocos_conciliados)
                    if blocos_pendentes:
                        dfs_ordenados.extend(blocos_pendentes)

                    df_reorganizado = pd.concat(dfs_ordenados, ignore_index=True) if dfs_ordenados else df_r
                    df_reorganizado["_id_raz_"] = range(len(df_reorganizado))
                    df_reorganizado["_id_excel_"] = range(2, len(df_reorganizado) + 2)

                    mapa_linhas_doc = {}
                    for doc, grp in df_reorganizado[df_reorganizado["_pre_conciliado_"]].groupby("_doc_extraido_"):
                        linhas_str = ", ".join(grp["_id_excel_"].astype(str).tolist())
                        mapa_linhas_doc[doc] = linhas_str

                    # ETAPA 3: AUDITORIA VIA IA EM LOTES INTELIGENTES
                    resultados_linhas = {}
                    TAMANHO_LOTE = 50
                    total_linhas = len(df_reorganizado)

                    barra_progresso = st.progress(0)
                    status_texto = st.empty()

                    i = 0
                    while i < total_linhas:
                        fim = min(i + TAMANHO_LOTE, total_linhas)
                        if fim < total_linhas:
                            doc_borda = df_reorganizado.iloc[fim - 1]["_doc_extraido_"]
                            while fim < total_linhas and df_reorganizado.iloc[fim]["_doc_extraido_"] == doc_borda:
                                fim += 1

                        lote_df = df_reorganizado.iloc[i:fim]
                        progresso_atual = min(1.0, fim / max(1, total_linhas))
                        barra_progresso.progress(progresso_atual)
                        status_texto.text(f"Agente auditando lote (linhas {i+2} a {fim+1} de {total_linhas+1})...")

                        csv_payload = lote_df[[
                            "_id_raz_", "_id_excel_", "_doc_extraido_",
                            "_val_deb_", "_val_cred_", "_pre_conciliado_", col_titulo_raz
                        ]].to_csv(index=False, sep=";")

                        prompt_sistema = (
                            "Você atua como um Agente Auditor Contábil Sênior.\n"
                            "Os dados estão delimitados por ponto e vírgula (;).\n"
                            "Regras:\n"
                            "1. Se '_pre_conciliado_' for True, classifique como 'CONCILIADO' explicando a compensação entre as linhas do Excel ('_id_excel_').\n"
                            "2. Se houver sobras reais sem contrapartida, classifique como divergente/em aberto.\n"
                            "3. NUNCA use ponto e vírgula (;) no texto da explicação.\n\n"
                            "Retorne estritamente em formato CSV delimitado por ponto e vírgula (;), com as colunas:\n"
                            "_id_raz_;status;explicacao\n"
                            "Retorne apenas o texto puro sem blocos markdown."
                        )

                        max_tentativas = 3
                        sucesso_lote = False
                        for tentativa in range(max_tentativas):
                            try:
                                response = client.chat.completions.create(
                                    model=modelo,
                                    messages=[
                                        {"role": "system", "content": prompt_sistema},
                                        {"role": "user", "content": csv_payload}
                                    ],
                                    temperature=0.0
                                )
                                resposta_csv = response.choices[0].message.content.strip()
                                resposta_csv = re.sub(r'```[a-zA-Z]*\n?', '', resposta_csv).replace('```', '').strip()

                                for linha in resposta_csv.splitlines():
                                    partes = linha.split(";", 2)
                                    if len(partes) == 3 and partes[0].strip().isdigit():
                                        rid = int(partes[0].strip())
                                        status_ia = partes[1].strip().upper()
                                        exp_ia = partes[2].strip()
                                        if status_ia in ["CONCILIADO", "DIVERGENTE"]:
                                            resultados_linhas[rid] = {
                                                "status": status_ia,
                                                "explicacao": exp_ia
                                            }
                                sucesso_lote = True
                                break
                            except Exception as err:
                                if "503" in str(err) or "429" in str(err):
                                    time.sleep(2 * (tentativa + 1))
                                else:
                                    break

                        i = fim

                    barra_progresso.empty()
                    status_texto.empty()

                    # ETAPA 4: CONSOLIDAÇÃO FINAL
                    lista_final = []
                    colunas_remover = [
                        "_id_raz_", "_id_excel_", "_val_deb_", "_val_cred_",
                        "_doc_extraido_", "_pre_conciliado_"
                    ]

                    for idx, row in df_reorganizado.iterrows():
                        rid = row["_id_raz_"]
                        doc = row["_doc_extraido_"]
                        pre_conc = row["_pre_conciliado_"]

                        res = resultados_linhas.get(rid)

                        if pre_conc:
                            linhas_par = mapa_linhas_doc.get(doc, str(row["_id_excel_"]))
                            if not res or res["status"] != "CONCILIADO":
                                res = {
                                    "status": "CONCILIADO",
                                    "explicacao": f"Conciliado: Compensação validada entre as linhas {linhas_par}."
                                }
                        elif not res:
                            res = {
                                "status": "DIVERGENTE",
                                "explicacao": "Divergente: Lançamento em aberto sem contrapartida correspondente."
                            }

                        row_dict = row.to_dict()
                        for col_rm in colunas_remover:
                            row_dict.pop(col_rm, None)

                        row_dict["status"] = res["status"]
                        row_dict["explicacao"] = res["explicacao"]
                        lista_final.append(row_dict)

                    df_razao_final = pd.DataFrame(lista_final)

                    st.success("🎉 Processamento, segunda mão de revisão e auditoria concluídos com sucesso!")
                    st.session_state["df_razao_final"] = df_razao_final
                    st.session_state["analise_concluida_reorganizada"] = True

                except Exception as e:
                    st.error(f"❌ Erro durante o processamento: {e}")

    # --- EXIBIÇÃO E DOWNLOAD ---
    if st.session_state.get("analise_concluida_reorganizada", False):
        st.markdown("---")
        st.subheader("📋 Planilha Reorganizada, Auditada e Downloads")

        df_r_f = st.session_state["df_razao_final"]
        st.dataframe(df_r_f, use_container_width=True)

        def gerar_excel_bytes(df_dados):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Razao_Conciliado"
            headers = list(df_dados.columns)
            ws.append(headers)

            verde_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
            verde_font = Font(color="155724", bold=True)

            for row_idx, row in df_dados.iterrows():
                ws.append([row[col] for col in headers])
                if str(row.get("status", "")) == "CONCILIADO":
                    current_row_excel = row_idx + 2
                    for col_num in range(1, len(headers) + 1):
                        cell = ws.cell(row=current_row_excel, column=col_num)
                        cell.fill = verde_fill
                        cell.font = verde_font
            wb.save(output)
            return output.getvalue()

        excel_bytes = gerar_excel_bytes(df_r_f)
        st.download_button(
            label="💾 Baixar Razão Reorganizado e Conciliado (Excel)",
            data=excel_bytes,
            file_name="razao_reorganizado_conciliado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Por favor, envie o arquivo do Razão para iniciar.")