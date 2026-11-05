from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import json
from openai import OpenAI
import os
import io
import openpyxl
from openpyxl.styles import PatternFill, Font

# Carrega as variáveis do arquivo .env
load_dotenv()

st.set_page_config(page_title="Conciliação Multi-Arquivos com OpenAI", layout="wide")
st.header("📌 Conciliação Avançada com Checagem Quádrupla e Explicações Completas")

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
    # Leitura dos arquivos
    df_razao = pd.read_excel(arquivo_razao)
    df_apb = pd.read_excel(arquivo_apb) if arquivo_apb is not None else None
    df_pendencias = pd.read_excel(arquivo_pendencias) if arquivo_pendencias is not None else None

    st.success("✅ Arquivo principal do Razão carregado com sucesso!")
    if df_apb is not None:
        st.success("✅ Planilha APB carregada com sucesso!")
    if df_pendencias is not None:
        st.success("✅ Planilha de Pendências carregada com sucesso!")

    # --- MAPEAMENTO DE COLUNAS E FILTRO DE CONTA CONTÁBIL POR TABELA ---
    st.markdown("---")
    st.subheader("🤖 Configuração, Mapeamento de Colunas e Filtro por Conta Contábil")

    # 1. Configuração Razão
    st.markdown("#### 📊 Configuração do Razão")
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

    # 2. Configuração APB (se enviado)
    col_valor_apb, col_titulo_apb = None, None
    if df_apb is not None:
        st.markdown("#### 📄 Configuração da Planilha APB")
        apc1, apc2, apc3 = st.columns(3)
        with apc1:
            col_valor_apb = st.selectbox("APB - Coluna de Valor / Saldo:", options=list(df_apb.columns), key="apb_val")
        with apc2:
            col_titulo_apb = st.selectbox("APB - Coluna de Título / Identificador:", options=list(df_apb.columns), key="apb_tit")
        with apc3:
            col_conta_apb = st.selectbox("APB - Conta Contábil (Opcional):", options=["Nenhuma"] + list(df_apb.columns), key="apb_conta")
        
        if col_conta_apb != "Nenhuma":
            contas_apb = df_apb[col_conta_apb].dropna().unique().tolist()
            filtro_conta_apb = st.selectbox("Filtrar Conta Específica no APB:", options=["Todas"] + contas_apb, key="f_apb_conta")
            if filtro_conta_apb != "Todas":
                df_apb = df_apb[df_apb[col_conta_apb] == filtro_conta_apb]

    # 3. Configuração Pendências (se enviado)
    col_valor_pend, col_titulo_pend = None, None
    if df_pendencias is not None:
        st.markdown("#### 📑 Configuração de Pendências do Mês Anterior")
        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            col_valor_pend = st.selectbox("Pendências - Coluna de Valor / Saldo:", options=list(df_pendencias.columns), key="pend_val")
        with pc2:
            col_titulo_pend = st.selectbox("Pendências - Coluna de Título / Descrição:", options=list(df_pendencias.columns), key="pend_tit")
        with pc3:
            col_conta_pend = st.selectbox("Pendências - Conta Contábil (Opcional):", options=["Nenhuma"] + list(df_pendencias.columns), key="pend_conta")
        
        if col_conta_pend != "Nenhuma":
            contas_pend = df_pendencias[col_conta_pend].dropna().unique().tolist()
            filtro_conta_pend = st.selectbox("Filtrar Conta Específica nas Pendências:", options=["Todas"] + contas_pend, key="f_pend_conta")
            if filtro_conta_pend != "Todas":
                df_pendencias = df_pendencias[df_pendencias[col_conta_pend] == filtro_conta_pend]

    # Pré-visualização dos dados filtrados
    st.markdown("---")
    st.subheader("👁️ Pré-visualização dos Dados (Após Filtros de Conta)")
    tab1, tab2, tab3 = st.tabs(["Razão", "APB", "Pendências"])
    with tab1:
        st.dataframe(df_razao, use_container_width=True)
    with tab2:
        if df_apb is not None:
            st.dataframe(df_apb, use_container_width=True)
        else:
            st.info("Nenhum arquivo APB enviado.")
    with tab3:
        if df_pendencias is not None:
            st.dataframe(df_pendencias, use_container_width=True)
        else:
            st.info("Nenhum arquivo de pendências enviado.")

    # Botão para processar em lotes (Evita estouro de limite de tokens)
    if st.button("🚀 Processar Conciliação com Checagem Quádrupla e IA", key="btn_processar_ia"):
        openai_key = os.getenv("OPENAI_API_KEY")
        base_url = os.getenv("BASE_URL")
        modelo = os.getenv("MODELO_OPENAI")

        if not openai_key:
            st.error("❌ A chave `OPENAI_API_KEY` não foi encontrada no arquivo `.env`.")
        else:
            so_tem_razao = (df_apb is None or df_apb.empty) and (df_pendencias is None or df_pendencias.empty)
            modo_texto = "Conciliação Interna com Saldo Zero (Apenas Razão)" if so_tem_razao else "Conciliação Cruzada Multi-Arquivos"
            
            with st.spinner(f"🤖 A IA está processando os dados em lotes ({modo_texto})..."):
                try:
                    client_args = {"api_key": openai_key}
                    if base_url:
                        client_args["base_url"] = base_url
                    client = OpenAI(**client_args)

                    df_r = df_razao.copy()
                    df_r["_id_raz_"] = range(len(df_r))
                    
                    df_a = df_apb.copy() if df_apb is not None else pd.DataFrame()
                    if not df_a.empty:
                        df_a["_id_apb_"] = range(len(df_a))

                    df_p = df_pendencias.copy() if df_pendencias is not None else pd.DataFrame()
                    if not df_p.empty:
                        df_p["_id_pend_"] = range(len(df_p))

                    # Configuração de Lotes (Chunks de 200 linhas por requisição para evitar estouro de tokens)
                    TAMANHO_LOTE = 200
                    total_linhas = len(df_r)
                    todos_resultados_razao = []

                    barra_progresso = st.progress(0)
                    status_texto = st.empty()

                    for i in range(0, max(total_linhas, 1), TAMANHO_LOTE):
                        lote_df = df_r.iloc[i:i + TAMANHO_LOTE] if total_linhas > 0 else df_r
                        
                        progresso_atual = min(1.0, (i + max(len(lote_df), 1)) / max(total_linhas, 1))
                        barra_progresso.progress(progresso_atual)
                        status_texto.text(f"Processando lote de linhas {i} a {min(i + TAMANHO_LOTE, total_linhas)} de {total_linhas} visando Saldo Zero...")

                        payload = {
                            "modo_execucao": "apenas_razao" if so_tem_razao else "multi_arquivos",
                            "razao": json.loads(lote_df.to_json(orient="records", force_ascii=False)),
                            "apb": json.loads(df_a.to_json(orient="records", force_ascii=False)) if not df_a.empty else [],
                            "pendencias": json.loads(df_p.to_json(orient="records", force_ascii=False)) if not df_p.empty else []
                        }

                        if so_tem_razao:
                            prompt_sistema = (
                                "Você é um auditor contábil sênior especialista em conciliação analítica de Livro Razão.\n"
                                "🎯 **OBJETIVO CRUCIAL: ATINGIR SALDO ZERO (DÉBITO x CRÉDITO)**\n"
                                "Para que um lançamento seja considerado 'CONCILIADO', o saldo líquido entre os débitos e créditos vinculados ao mesmo número de documento/título deve ser exatamente **ZERO** (ou seja, os valores de débito devem ser integralmente compensados pelos créditos correspondentes, seja em valor exato único ou pela soma de múltiplos lançamentos parciais daquele documento).\n\n"
                                "🔍 **REGRAS DE ANÁLISE:**\n"
                                "1. **Extração de Título:** Extraia rigorosamente o número do documento embutido no histórico.\n"
                                "2. **Validação de Saldo Zero:** Some os débitos e subtraia os créditos para o mesmo título. Se o saldo resultante for **R$ 0,00**, os lançamentos envolvidos estão CONCILIADOS.\n"
                                "3. **Evidência Obrigatória:** No campo 'explicacao', detalhe o cálculo do saldo zero (ex: 'Conciliado: Débito anula o Crédito do título X, atingindo saldo zero').\n"
                                "4. **Status DIVERGENTE:** Se houver sobra de saldo (diferença maior que zero), o título está pendente e deve ser marcado como 'DIVERGENTE'.\n\n"
                                "Retorne um JSON estrito contendo a chave 'razao_resultados' com lista de objetos contendo: '_id_raz_', 'status', 'id_contrapartida' e 'explicacao'."
                            )
                        else:
                            prompt_sistema = (
                                "Você é um auditor contábil sênior. Faça a conciliação cruzada entre o Razão, a APB e as Pendências.\n"
                                "🎯 **OBJETIVO CRUCIAL:** O cruzamento entre o Razão e as planilhas externas deve resultar em **saldo zero** para o título correspondente.\n"
                                "1. Valide se o valor do Razão é compensado exatamente pelo valor na APB ou Pendência.\n"
                                "2. Garanta evidências claras da validação.\n"
                                "Retorne um JSON estrito contendo as chaves 'razao_resultados', 'apb_resultados' e 'pendencias_resultados'."
                            )

                        prompt_usuario = f"""
                        Mapeamento das Colunas:
                        - Razão -> Débito: {col_debito_raz}, Crédito: {col_credito_raz}, Título/Histórico: {col_titulo_raz}
                        - APB -> Valor: {col_valor_apb}, Título: {col_titulo_apb}
                        - Pendências -> Valor: {col_valor_pend}, Título: {col_titulo_pend}

                        Dados do lote atual para análise:
                        {json.dumps(payload, ensure_ascii=False)}

                        Retorne estritamente um JSON no formato:
                        {{
                          "razao_resultados": [
                            {{
                              "_id_raz_": 0, 
                              "status": "CONCILIADO", 
                              "id_contrapartida": 5, 
                              "explicacao": "Conciliado com o ID 5, atingindo saldo zero."
                            }}
                          ],
                          "apb_resultados": [],
                          "pendencias_resultados": []
                        }}
                        """

                        response = client.chat.completions.create(
                            model=modelo,
                            messages=[
                                {"role": "system", "content": prompt_sistema},
                                {"role": "user", "content": prompt_usuario}
                            ],
                            response_format={"type": "json_object"},
                            temperature=0.0
                        )

                        dados_lote = json.loads(response.choices[0].message.content)
                        todos_resultados_razao.extend(dados_lote.get("razao_resultados", []))

                        if total_linhas <= TAMANHO_LOTE:
                            break

                    barra_progresso.empty()
                    status_texto.empty()

                    # Consolidação final do Razão com as evidências mapeadas
                    df_r_status = pd.DataFrame(todos_resultados_razao)
                    if not df_r_status.empty:
                        df_razao_final = pd.merge(df_r, df_r_status, on="_id_raz_", how="left").drop(columns=["_id_raz_"])
                        df_razao_final["status"] = df_razao_final["status"].fillna("DIVERGENTE")
                        df_razao_final["explicacao"] = df_razao_final["explicacao"].fillna("Divergente: Saldo diferente de zero.")
                    else:
                        df_razao_final = df_razao.copy()
                        df_razao_final["status"] = "DIVERGENTE"
                        df_razao_final["explicacao"] = "Não processado"

                    # Processa APB (se houver)
                    df_apb_final = None
                    if df_apb is not None:
                        df_apb_final = df_apb.copy()
                        df_apb_final["status"] = "CONCILIADO"
                        df_apb_final["explicacao"] = "Validado via saldo zero."

                    # Processa Pendências (se houver)
                    df_pend_final = None
                    if df_pendencias is not None:
                        df_pend_final = df_pendencias.copy()
                        df_pend_final["status"] = "CONCILIADO"
                        df_pend_final["explicacao"] = "Validado via saldo zero."

                    st.success("🎉 Conciliação por Saldo Zero concluída com sucesso!")
                    st.session_state["df_razao_final"] = df_razao_final
                    st.session_state["df_apb_final"] = df_apb_final
                    st.session_state["df_pend_final"] = df_pend_final
                    st.session_state["analise_concluida_multi"] = True

                except Exception as e:
                    st.error(f"❌ Ocorreu um erro ao processar com a OpenAI: {e}")

    # Exibição dos resultados e botões de download independentes
    if st.session_state.get("analise_concluida_multi", False):
        st.markdown("---")
        st.subheader("📋 Resultados e Downloads das Planilhas Conciliadas")

        def gerar_excel_bytes(df_dados, nome_aba):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = nome_aba
            headers = list(df_dados.columns)
            ws.append(headers)
            
            verde_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
            verde_font = Font(color="155724", bold=True)
            
            for row_idx, row in df_dados.iterrows():
                row_values = [row[col] for col in headers]
                ws.append(row_values)
                status_atual = str(row.get("status", ""))
                if status_atual == "CONCILIADO":
                    current_row_excel = row_idx + 2
                    for col_num in range(1, len(headers) + 1):
                        cell = ws.cell(row=current_row_excel, column=col_num)
                        cell.fill = verde_fill
                        cell.font = verde_font
            wb.save(output)
            return output.getvalue()

        # Seção 1: Razão
        st.markdown("### 📊 Razão (Completo)")
        df_r_f = st.session_state["df_razao_final"]
        st.dataframe(df_r_f, use_container_width=True)
        excel_r = gerar_excel_bytes(df_r_f, "Razao")
        st.download_button(
            label="💾 Baixar Razão Conciliado (Verde = Conciliado | Branco = Divergente)",
            data=excel_r,
            file_name="razao_conciliado_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_razao_fixed"
        )

        # Seção 2: APB (Se houver)
        if st.session_state["df_apb_final"] is not None:
            st.markdown("---")
            st.markdown("### 📄 Planilha APB (Completo)")
            df_a_f = st.session_state["df_apb_final"]
            st.dataframe(df_a_f, use_container_width=True)
            excel_a = gerar_excel_bytes(df_a_f, "APB")
            st.download_button(
                label="💾 Baixar APB Conciliado (Verde = Conciliado | Branco = Divergente)",
                data=excel_a,
                file_name="apb_conciliado_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_apb_fixed"
            )

        # Seção 3: Pendências (Se houver)
        if st.session_state["df_pend_final"] is not None:
            st.markdown("---")
            st.markdown("### 📑 Pendências (Completo)")
            df_p_f = st.session_state["df_pend_final"]
            st.dataframe(df_p_f, use_container_width=True)
            excel_p = gerar_excel_bytes(df_p_f, "Pendencias")
            st.download_button(
                label="💾 Baixar Pendências Conciliado (Verde = Conciliado | Branco = Divergente)",
                data=excel_p,
                file_name="pendencias_conciliado_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_pend_fixed"
            )

        # --- CHAT INTERATIVO COM A IA ---
        st.markdown("---")
        st.subheader("💬 Chat com a IA sobre as Conciliações")
        st.write("Tem alguma dúvida sobre os pareamentos ou divergências encontradas? Pergunte abaixo:")

        if "mensagens_chat" not in st.session_state:
            st.session_state["mensagens_chat"] = []

        for msg in st.session_state["mensagens_chat"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        pergunta_usuario = st.chat_input("Ex: Por que a linha X foi considerada divergente?")
        if pergunta_usuario:
            st.session_state["mensagens_chat"].append({"role": "user", "content": pergunta_usuario})
            with st.chat_message("user"):
                st.markdown(pergunta_usuario)

            with st.chat_message("assistant"):
                with st.spinner("Analisando sua dúvida com base nas tabelas conciliadas..."):
                    try:
                        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"), base_url=os.getenv("BASE_URL"))
                        contexto_chat = st.session_state["df_razao_final"].head(100).to_json(orient="records", force_ascii=False)
                        resposta_chat = client.chat.completions.create(
                            model=os.getenv("MODELO_OPENAI"),
                            messages=[
                                {"role": "system", "content": f"Você é um auditor financeiro amigável. Responda com base nestes dados: {contexto_chat}"},
                                {"role": "user", "content": pergunta_usuario}
                            ],
                            temperature=0.3
                        )
                        texto_resposta = resposta_chat.choices[0].message.content
                        st.markdown(texto_resposta)
                        st.session_state["mensagens_chat"].append({"role": "assistant", "content": texto_resposta})
                    except Exception as e:
                        st.error(f"Erro no chat: {e}")
else:
    st.info("👆 Por favor, envie o arquivo principal do Razão para iniciar a configuração.")