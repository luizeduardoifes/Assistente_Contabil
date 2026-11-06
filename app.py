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
st.header("🤖 Agente Auditor IA: Conciliação Cirúrgica via CSV e Lotes Inteligentes")

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

    # --- PROCESSAMENTO COM O AGENTE IA USANDO FORMATO CSV E LOTES DE SEGURANÇA ---
    if st.button("🚀 Executar Agente Auditor Cirúrgico (IA + CSV)", key="btn_processar_ia_cirurgico"):
        openai_key = os.getenv("OPENAI_API_KEY")
        base_url = os.getenv("BASE_URL")
        modelo = os.getenv("MODELO_OPENAI", "gpt-4o-mini")

        if not openai_key:
            st.error("❌ A chave `OPENAI_API_KEY` não foi encontrada no arquivo `.env`.")
        else:
            with st.spinner("🤖 O Agente Auditor está analisando a contabilidade de forma cirúrgica..."):
                try:
                    client_args = {"api_key": openai_key}
                    if base_url:
                        client_args["base_url"] = base_url
                    client = OpenAI(**client_args)

                    df_r = df_razao.copy()
                    df_r["_id_raz_"] = range(len(df_r))
                    df_r["_id_excel_"] = range(2, len(df_r) + 2) # Referência da linha real no Excel

                    # Tratamento numérico seguro
                    def parse_val(v):
                        if pd.isna(v) or str(v).strip() in ["-", "", " - "]:
                            return 0.0
                        s = str(v).replace(" ", "").replace(".", "").replace(",", ".")
                        try:
                            return float(s)
                        except:
                            return 0.0

                    df_r["_val_deb_"] = df_r[col_debito_raz].apply(parse_val)
                    df_r["_val_cred_"] = df_r[col_credito_raz].apply(parse_val)

                    resultados_linhas = {}

                    # Estratégia de divisão em lotes de 50 linhas para garantir estabilidade e precisão cirúrgica
                    TAMANHO_LOTE = 50
                    total_linhas = len(df_r)
                    
                    barra_progresso = st.progress(0)
                    status_texto = st.empty()

                    for i in range(0, total_linhas, TAMANHO_LOTE):
                        lote_df = df_r.iloc[i:i + TAMANHO_LOTE]
                        progresso_atual = min(1.0, (i + len(lote_df)) / total_linhas)
                        barra_progresso.progress(progresso_atual)
                        status_texto.text(f"Agente analisando lote contábil (linhas {i+2} a {i+len(lote_df)+1} de {total_linhas+1})...")

                        # Converte o lote atual para CSV limpo (otimizando tokens para a IA)
                        csv_payload = lote_df[["_id_raz_", "_id_excel_", col_debito_raz, col_credito_raz, col_titulo_raz]].to_csv(index=False)

                        # PROMPT DO SISTEMA COM O PASSO A PASSO CONTÁBIL DETALHADO
                        prompt_sistema = (
                            "Você atua como um Agente Auditor Contábil Sênior especializado em conciliação de Razão.\n"
                            "Sua missão é realizar a conciliação cirúrgica de forma rigorosa seguindo estritamente estas etapas:\n"
                            "1. **Isolamento de Chaves e Documentos:** Identifique números de notas fiscais, contratos, borderôs ou chaves contábeis (ex: NR, RC, DP, BORD) no texto do histórico.\n"
                            "2. **Checagem de Partidas Dobradas e Saldo Zero:** Analise os valores das colunas de Débito e Crédito. Se houver correspondência entre uma implantação (crédito) e suas respectivas baixas (débito) cujo somatório zere perfeitamente (Saldo Zero = 0.00), classifique o status como 'CONCILIADO'.\n"
                            "3. **Parecer Detalhado:** No campo 'explicacao', escreva um laudo técnico explicando com clareza a conciliação e citando os números das linhas Excel envolvidas.\n"
                            "4. Caso não haja contrapartida correspondente que zere o saldo, classifique como 'DIVERGENTE'.\n\n"
                            "Retorne sua resposta estritamente no formato CSV delimitado por ponto e vírgula (;), contendo exatamente as colunas:\n"
                            "_id_raz_;status;explicacao\n"
                            "Não inclua nenhum bloco de markdown extra (como ```csv), retorne apenas o texto delimitado."
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

                                df_resp = pd.read_csv(io.StringIO(resposta_csv), sep=";")
                                for _, r_item in df_resp.iterrows():
                                    try:
                                        rid = int(r_item["_id_raz_"])
                                        resultados_linhas[rid] = {
                                            "status": str(r_item["status"]).strip().upper(),
                                            "explicacao": str(r_item["explicacao"]).strip()
                                        }
                                    except:
                                        pass
                                sucesso_lote = True
                                break
                            except Exception as err:
                                if "503" in str(err) or "429" in str(err):
                                    time.sleep(2 * (tentativa + 1))
                                else:
                                    break

                        if not sucesso_lote:
                            for _, row_lote in lote_df.iterrows():
                                rid = row_lote["_id_raz_"]
                                if rid not in resultados_linhas:
                                    resultados_linhas[rid] = {
                                        "status": "DIVERGENTE",
                                        "explicacao": "Divergente: Sem contrapartida validada no lote."
                                    }

                    barra_progresso.empty()
                    status_texto.empty()

                    # Consolidação final do DataFrame
                    lista_final = []
                    for idx, row in df_r.iterrows():
                        rid = row["_id_raz_"]
                        res = resultados_linhas.get(rid, {"status": "DIVERGENTE", "explicacao": "Divergente: Lançamento sem correspondência."})
                        row_dict = row.to_dict()
                        row_dict.pop("_id_raz_", None)
                        row_dict.pop("_id_excel_", None)
                        row_dict.pop("_val_deb_", None)
                        row_dict.pop("_val_cred_", None)
                        
                        row_dict["status"] = res["status"]
                        row_dict["explicacao"] = res["explicacao"]
                        lista_final.append(row_dict)

                    df_razao_final = pd.DataFrame(lista_final)

                    st.success("🎉 Conciliação cirúrgica com IA concluída com sucesso!")
                    st.session_state["df_razao_final"] = df_razao_final
                    st.session_state["analise_concluida_cirurgica"] = True

                except Exception as e:
                    st.error(f"❌ Erro durante o processamento com o Agente: {e}")

    # --- EXIBIÇÃO E DOWNLOAD ---
    if st.session_state.get("analise_concluida_cirurgica", False):
        st.markdown("---")
        st.subheader("📋 Resultados da Auditoria Cirúrgica e Downloads")

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
            label="💾 Baixar Razão Conciliado Cirúrgico (Excel Formatado)",
            data=excel_bytes,
            file_name="razao_conciliado_cirurgico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Por favor, envie o arquivo do Razão para iniciar a auditoria cirúrgica com IA.")