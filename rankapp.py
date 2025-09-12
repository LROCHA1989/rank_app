import streamlit as st
import pandas as pd
import locale
import streamlit as st
from io import BytesIO


# 🧠 Configurar moeda brasileira
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')

# 🔧 Função para limpar valores monetários
def limpar_valor(valor):
    if isinstance(valor, str):
        valor = valor.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return float(valor)
    except:
        return None

# 🗂️ Inicializar sessão
if "dados_vendas" not in st.session_state:
    st.session_state["dados_vendas"] = pd.DataFrame()

# 🎛️ Menu lateral
opcao = st.sidebar.radio("📌 NAVEGAÇÃO!", [
    "📤 Upload de Arquivo",
    "📊 Venda Geral",
    "🏆 Classificação Geral",
    "📈 Análise de Variação Anual"
])

# 📤 Upload de Arquivo
# 📤 Upload de Arquivo
if opcao == "📤 Upload de Arquivo":
    st.title("📤 Upload de Arquivo Excel")
    st.write("Envie um arquivo `.xlsx` com as abas `VENDAS` e `PONTOS_EXTRAS`.")

    arquivo = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"], key="upload")

    if arquivo:
        try:
            xls = pd.ExcelFile(arquivo)

            # 🟢 Ler aba de vendas
            df = pd.read_excel(xls, sheet_name="VENDAS")

            # 🟢 Ler aba de pontos extras (se existir)
            if "PONTOS_EXTRAS" in xls.sheet_names:
                pontos = pd.read_excel(xls, sheet_name="PONTOS_EXTRAS")
            else:
                pontos = pd.DataFrame(columns=["REP.", "MÊS", "AÇÃO", "PROMOÇÃO", "INADIMPLÊNCIA"])

            # Padronizar e limpar colunas de vendas
            df["MÊS"] = df["MÊS"].astype(str).str.upper().str.strip().str.replace(".", "", regex=False)
            meses_corretos = {
                "JAN": "JAN", "FEV": "FEV", "MAR": "MAR", "ABR": "ABR", "MAI": "MAI", "JUN": "JUN",
                "JUL": "JUL", "AGO": "AGO", "SET": "SET", "OUT": "OUT", "NOV": "NOV", "DEZ": "DEZ",
                "VEF": "FEV", "DEFINIR": "SET", "ATRAS": "AGO", "FEB": "FEV", "SEPT": "SET", "SEP": "SET", "DEC": "DEZ",
                "JANEIRO": "JAN", "FEVEREIRO": "FEV", "MARÇO": "MAR", "ABRIL": "ABR", "MAIO": "MAI",
                "JUNHO": "JUN", "JULHO": "JUL", "AGOSTO": "AGO", "SETEMBRO": "SET", "OUTUBRO": "OUT",
                "NOVEMBRO": "NOV", "DEZEMBRO": "DEZ"
            }
            df["MÊS"] = df["MÊS"].map(meses_corretos).fillna(df["MÊS"])
            df["REP."] = df["REP."].astype(str).str.upper().str.strip()
            df["EMPRESA"] = df["EMPRESA"].astype(str).str.upper().str.strip()
            df["ANO"] = pd.to_numeric(df["ANO"], errors="coerce", downcast="integer")
            df["SUBTOTAL"] = df["SUBTOTAL"].apply(limpar_valor)

            # Padronizar pontos extras
            pontos["REP."] = pontos["REP."].astype(str).str.upper().str.strip()
            pontos["MÊS"] = pontos["MÊS"].astype(str).str.upper().str.strip()
            pontos["AÇÃO"] = pd.to_numeric(pontos["AÇÃO"], errors="coerce").fillna(0).astype(int)
            pontos["PROMOÇÃO"] = pd.to_numeric(pontos["PROMOÇÃO"], errors="coerce").fillna(0).astype(int)
            pontos["INADIMPLÊNCIA"] = pd.to_numeric(pontos["INADIMPLÊNCIA"], errors="coerce").fillna(0).astype(int)

            # Salvar nas variáveis de sessão
            st.session_state["dados_vendas"] = df.copy()
            st.session_state["pontos_extras"] = pontos.copy()

            st.success("Arquivo carregado com sucesso!")

        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")

    # Exibir dados carregados e botão de limpar
    if not st.session_state["dados_vendas"].empty:
        if st.button("🗑️ Limpar dados carregados"):
            st.session_state["dados_vendas"] = pd.DataFrame()
            st.session_state["pontos_extras"] = pd.DataFrame(columns=["REP.", "MÊS", "AÇÃO", "PROMOÇÃO", "INADIMPLÊNCIA"])
            st.success("Dados removidos com sucesso!")
            st.stop()

        with st.expander("📄 Visualizar dados carregados"):
            st.dataframe(st.session_state["dados_vendas"][["REP.", "SUBTOTAL", "MÊS", "EMPRESA", "ANO"]])

        with st.expander("📌 Visualizar pontos extras"):
            st.dataframe(st.session_state["pontos_extras"])

# 📊 Venda Geral
elif opcao == "📊 Venda Geral":
    st.title("📊 VENDA GERAL")

    df = st.session_state["dados_vendas"]

    if df.empty:
        st.warning("Nenhum dado disponível. Faça o upload de um arquivo primeiro.")
    else:
        # Limpar valores
        df["SUBTOTAL"] = df["SUBTOTAL"].apply(limpar_valor)
        df = df[df["SUBTOTAL"].notnull() & (df["SUBTOTAL"] > 0)]

        # Filtro por ano
        anos_disponiveis = sorted(df["ANO"].dropna().unique())
        ano_selecionado = st.selectbox("Filtrar por ano", options=anos_disponiveis, index=len(anos_disponiveis)-1)


        df_filtrado = df[df["ANO"] == ano_selecionado].copy()
        df_filtrado["MÊS"] = df_filtrado["MÊS"].str[:3].str.upper()

        # Ordem fixa dos meses
        ordem_meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
                       "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

        # Criar tabela dinâmica
        tabela_dinamica = pd.pivot_table(
            df_filtrado,
            index="REP.",
            columns="MÊS",
            values="SUBTOTAL",
            aggfunc="sum",
            fill_value=0
        )

        tabela_dinamica = tabela_dinamica.reindex(columns=ordem_meses, fill_value=0)
        tabela_dinamica["TOTAL GERAL"] = tabela_dinamica.sum(axis=1)

        # Adicionar linha TOTAL POR MÊS
        total_por_mes = tabela_dinamica.sum(axis=0)
        tabela_dinamica.loc["TOTAL POR MÊS"] = total_por_mes

        # Separar linha de total
        linha_total = tabela_dinamica.loc[["TOTAL POR MÊS"]]
        tabela_dinamica_sem_total = tabela_dinamica.drop(index="TOTAL POR MÊS")

        # Ordenar por TOTAL GERAL
        tabela_ordenada = tabela_dinamica_sem_total.sort_values(by="TOTAL GERAL", ascending=False)
        tabela_final = pd.concat([tabela_ordenada, linha_total])

        # Estilo visual
        def destacar_total(val):
            return ["font-weight: bold; background-color: #f0f0f0" if val.name == "TOTAL POR MÊS" else "" for _ in val]

        tabela_formatada = (
            tabela_final.style
            .format(lambda x: locale.currency(x, grouping=True))
            .apply(destacar_total, axis=1)
        )

        st.subheader(f"📋 Vendas por Mês - Ano {ano_selecionado}")
        st.dataframe(tabela_formatada)

        # ✅ Narrativa de desempenho por representante
        st.subheader("🗣️ Narrativa de Representantes")

        # Ranking geral no ano selecionado
        ranking_ano = df_filtrado.groupby("REP.")["SUBTOTAL"].sum().sort_values(ascending=False)
        melhores = ranking_ano.head(5)
        piores = ranking_ano.tail(5)

        # Variação entre 2025 e 2024
        if 2025 in anos_disponiveis and 2024 in anos_disponiveis:
            vendas_2025 = df[df["ANO"] == 2025].groupby("REP.")["SUBTOTAL"].sum()
            vendas_2024 = df[df["ANO"] == 2024].groupby("REP.")["SUBTOTAL"].sum()

            comparativo = pd.DataFrame({
                "2024": vendas_2024,
                "2025": vendas_2025
            }).fillna(0)

            comparativo["VARIAÇÃO (%)"] = ((comparativo["2025"] - comparativo["2024"]) /
                                           (comparativo["2024"].replace(0, 1))) * 100

            variacoes = comparativo["VARIAÇÃO (%)"].sort_values(ascending=False)
            top_crescimento = variacoes.head(5)
            top_queda = variacoes.tail(5)

            narrativa = f"""
            <p style='font-size:16px'>
            No ano de <strong>{ano_selecionado}</strong>, os 5 representantes com maior volume de vendas foram:<br>
            <strong>{', '.join(melhores.index)}</strong>.<br><br>
            Os 5 com menor desempenho foram:<br>
            <strong>{', '.join(piores.index)}</strong>.<br><br>
            Comparando 2025 com 2024:<br>
            Os maiores crescimentos foram de <strong>{', '.join(top_crescimento.index)}</strong> — destaque para <strong>{top_crescimento.index[0]}</strong> com crescimento de <span style='color:green; font-weight:bold'>{top_crescimento.iloc[0]:.2f}%</span>.<br>
            As maiores quedas foram de <strong>{', '.join(top_queda.index)}</strong> — destaque para <strong>{top_queda.index[0]}</strong> com queda de <span style='color:red; font-weight:bold'>{top_queda.iloc[0]:.2f}%</span>.
            </p>
            """
            st.markdown(narrativa, unsafe_allow_html=True)
        else:
            st.info("Para gerar a narrativa de variação, é necessário que o arquivo contenha dados de 2024 e 2025.")

# 🏆 Classificação Geral
elif opcao == "🏆 Classificação Geral":
    st.title("🏆 Ranking Geral de Representantes")

    df = st.session_state["dados_vendas"]

    if df.empty:
        st.warning("Nenhum dado disponível. Faça o upload de um arquivo primeiro.")
    else:
        # ✅ Corrigir nomes dos meses para manter formato do Excel
        mes_map = {
            "jan": "JANEIRO", "fev": "FEVEREIRO", "mar": "MARÇO", "abr": "ABRIL",
            "mai": "MAIO", "jun": "JUNHO", "jul": "JULHO", "ago": "AGOSTO",
            "set": "SETEMBRO", "out": "OUTUBRO", "nov": "NOVEMBRO", "dez": "DEZEMBRO"
        }
        df["MÊS"] = df["MÊS"].astype(str).str.strip().str.lower().map(mes_map).fillna(df["MÊS"])

        df["SUBTOTAL"] = df["SUBTOTAL"].apply(limpar_valor)
        df = df[df["SUBTOTAL"].notnull() & (df["SUBTOTAL"] > 0)]

        # 🎛️ Filtros
        anos_disponiveis = sorted(df["ANO"].dropna().unique())
        ano_selecionado = st.selectbox("Filtrar por ano", options=anos_disponiveis, index=len(anos_disponiveis)-1)

        empresas_disponiveis = sorted(df["EMPRESA"].dropna().unique())
        empresa_selecionada = st.selectbox("Filtrar por empresa", options=["Todas"] + empresas_disponiveis)

        meses_disponiveis = sorted(df["MÊS"].dropna().unique())
        meses_selecionados = st.multiselect("Filtrar por mês", options=meses_disponiveis)

        # Aplicar filtros
        df_filtrado = df[df["ANO"] == ano_selecionado]
        if empresa_selecionada != "Todas":
            df_filtrado = df_filtrado[df_filtrado["EMPRESA"] == empresa_selecionada]
        if meses_selecionados:
            df_filtrado = df_filtrado[df_filtrado["MÊS"].isin(meses_selecionados)]

        # 🏆 Ranking com pontos
        ranking = df_filtrado.groupby("REP.")["SUBTOTAL"].sum().sort_values(ascending=False).reset_index()
        if ranking.empty:
            st.warning("⚠️ Mês sem dados de venda.")
            st.stop()
        # Calcular pontos com base na posição
        multiplicadores = [5, 4, 3, 2] + [1] * (len(ranking) - 4)
        ranking["PONTOS"] = (ranking["SUBTOTAL"] / 20000 * pd.Series(multiplicadores)).round().astype(int)

        # Formatando SUBTOTAL como moeda brasileira
        ranking["SUBTOTAL"] = ranking["SUBTOTAL"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))

        # 🔧 Campos para adicionar ou desfazer pontos extras
        st.markdown("### ➕ Gerenciar Pontos Extras por Representante")

        rep_input = st.text_input("Nome do Representante").strip()
        mes_input = st.selectbox("Mês da Pontuação", options=meses_disponiveis)
        acao_pontos = st.number_input("Pontos por Ação", min_value=0, step=1)
        promo_pontos = st.number_input("Pontos por Promoção", min_value=0, step=1)
        inad_pontos = st.number_input("Pontos por Inadimplência", min_value=0, step=1)

        # Criar ou recuperar DataFrame de pontos extras acumulativos
        if "pontos_extras" not in st.session_state:
            st.session_state["pontos_extras"] = pd.DataFrame(columns=["REP.", "MÊS", "AÇÃO", "PROMOÇÃO", "INADIMPLÊNCIA"])

        # Botões de ação
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Incluir Pontos"):
                if rep_input and mes_input:
                    existente = st.session_state["pontos_extras"][
                        (st.session_state["pontos_extras"]["REP."] == rep_input) &
                        (st.session_state["pontos_extras"]["MÊS"] == mes_input)
                    ]
                    if not existente.empty:
                        idx = existente.index[0]
                        st.session_state["pontos_extras"].loc[idx, "AÇÃO"] += acao_pontos
                        st.session_state["pontos_extras"].loc[idx, "PROMOÇÃO"] += promo_pontos
                        st.session_state["pontos_extras"].loc[idx, "INADIMPLÊNCIA"] += inad_pontos
                    else:
                        nova_linha = pd.DataFrame([{
                            "REP.": rep_input,
                            "MÊS": mes_input,
                            "AÇÃO": acao_pontos,
                            "PROMOÇÃO": promo_pontos,
                            "INADIMPLÊNCIA": inad_pontos
                        }])
                        st.session_state["pontos_extras"] = pd.concat([st.session_state["pontos_extras"], nova_linha], ignore_index=True)

        with col2:
            if st.button("❌ Desfazer Pontos"):
                if rep_input and mes_input:
                    st.session_state["pontos_extras"] = st.session_state["pontos_extras"][
                        ~((st.session_state["pontos_extras"]["REP."] == rep_input) &
                          (st.session_state["pontos_extras"]["MÊS"] == mes_input))
                    ]
                    st.success(f"Pontos removidos para {rep_input} no mês {mes_input}")

        # 📁 Exportar histórico de pontos extras
        st.markdown("### 📤 Exportar Histórico de Pontos Extras")
        if not st.session_state["pontos_extras"].empty:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                st.session_state["pontos_extras"].to_excel(writer, index=False, sheet_name="Histórico de Pontos")
            st.download_button(
                label="📥 Baixar Histórico em Excel",
                data=buffer.getvalue(),
                file_name="historico_pontos_extras.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Nenhum ponto extra registrado ainda.")
        
        # Filtrar pontos extras para os meses selecionados
        pontos_mes = st.session_state["pontos_extras"]
        if meses_selecionados:
            pontos_mes = pontos_mes[pontos_mes["MÊS"].isin(meses_selecionados)]

        # Inicializar colunas extras
        ranking["AÇÃO"] = 0
        ranking["PROMOÇÃO"] = 0
        ranking["INADIMPLÊNCIA"] = 0

        # Aplicar pontos extras por representante
        for _, linha in pontos_mes.iterrows():
            idx = ranking[ranking["REP."] == linha["REP."]].index
            if not idx.empty:
                ranking.loc[idx, "AÇÃO"] += linha["AÇÃO"]
                ranking.loc[idx, "PROMOÇÃO"] += linha["PROMOÇÃO"]
                ranking.loc[idx, "INADIMPLÊNCIA"] += linha["INADIMPLÊNCIA"]

        # Calcular total de pontos
        ranking["TOTAL DE PONTOS"] = ranking["PONTOS"] + ranking["AÇÃO"] + ranking["PROMOÇÃO"] + ranking["INADIMPLÊNCIA"]

        # Calcular totais simples por coluna
        total_subtotal_valor = df_filtrado["SUBTOTAL"].sum()
        total_pontos = ranking["PONTOS"].sum()
        total_acao = ranking["AÇÃO"].sum()
        total_promocao = ranking["PROMOÇÃO"].sum()
        total_inad = ranking["INADIMPLÊNCIA"].sum()
        total_geral = ranking["TOTAL DE PONTOS"].sum()

        linha_total = pd.DataFrame({
            "REP.": ["TOTAL GERAL"],
            "SUBTOTAL": [f"R$ {total_subtotal_valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")],
            "PONTOS": [total_pontos],
            "AÇÃO": [total_acao],
            "PROMOÇÃO": [total_promocao],
            "INADIMPLÊNCIA": [total_inad],
            "TOTAL DE PONTOS": [total_geral]
        })

        # Concatenar ranking com linha de total
        ranking_final = pd.concat([ranking, linha_total], ignore_index=True)

        # Adicionar coluna de posição
        medalhas = ["🥇", "🥈", "🥉", "🏅", "🏅"]
        posicoes = [medalhas[i] if i < 5 else f"{i+1}º" for i in range(len(ranking))]
        posicoes.append("🔢")  # Para TOTAL GERAL
        ranking_final.insert(0, "POSIÇÃO", posicoes)

        # 📋 Exibir tabela
        titulo = f"🏅 Classificação Geral - Ano {ano_selecionado}"
        if empresa_selecionada != "Todas":
            titulo += f" - {empresa_selecionada}"
        if meses_selecionados:
            titulo += " - Mês " + ", ".join(meses_selecionados)

        st.subheader(titulo)
        st.dataframe(ranking_final, use_container_width=True, hide_index=True)
                # 📥 Exportar Tabela de Classificação Geral (logo abaixo do histórico)
        st.markdown("### 📥 Exportar Tabela de Classificação Geral")
        buffer_tabela = BytesIO()
        with pd.ExcelWriter(buffer_tabela, engine="xlsxwriter") as writer:
            ranking_final.to_excel(writer, index=False, sheet_name="Classificação Geral")
        st.download_button(
            label="📥 Baixar Tabela de Classificação",
            data=buffer_tabela.getvalue(),
            file_name="classificacao_geral.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# 📈 Análise de Variação Anual
elif opcao == "📈 Análise de Variação Anual":
    st.title("📈 Análise de Variação entre 2024 e 2025")

    df = st.session_state["dados_vendas"]

    if df.empty or "ANO" not in df.columns:
        st.warning("Dados insuficientes. Certifique-se de que o arquivo contém a coluna 'ANO'.")
    else:
        # Limpar e preparar dados
        df["SUBTOTAL"] = df["SUBTOTAL"].apply(limpar_valor)
        df = df[df["SUBTOTAL"].notnull() & (df["SUBTOTAL"] > 0)]
        df["MÊS"] = df["MÊS"].str[:3].str.upper()

        # Agrupar por mês e ano
        vendas_2024 = df[df["ANO"] == 2024].groupby("MÊS")["SUBTOTAL"].sum()
        vendas_2025 = df[df["ANO"] == 2025].groupby("MÊS")["SUBTOTAL"].sum()

        # Criar DataFrame comparativo
        comparativo = pd.DataFrame({
            "2024": vendas_2024,
            "2025": vendas_2025
        }).fillna(0)

        # ✅ Reordenar meses na ordem cronológica
        ordem_meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
                       "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
        comparativo = comparativo.reindex(ordem_meses)

        # ✅ Adicionar coluna TOTAL GERAL por mês
        comparativo["TOTAL GERAL"] = comparativo["2024"] + comparativo["2025"]

        # Calcular variação percentual por mês
        comparativo["VARIAÇÃO (%)"] = ((comparativo["2025"] - comparativo["2024"]) /
                                       (comparativo["2024"].replace(0, 1))) * 100

        # ✅ Calcular variação total anual separadamente
        soma_2024 = comparativo.loc[ordem_meses, "2024"].sum()
        soma_2025 = comparativo.loc[ordem_meses, "2025"].sum()
        variacao_total = (soma_2025 - soma_2024) / (soma_2024 if soma_2024 != 0 else 1) * 100

        # ✅ Adicionar linha TOTAL GERAL no final
        linha_total = pd.DataFrame(comparativo.loc[ordem_meses].sum(numeric_only=True)).T
        linha_total.index = ["TOTAL GERAL"]
        linha_total["VARIAÇÃO (%)"] = variacao_total
        comparativo = pd.concat([comparativo, linha_total])

        # Identificar melhores e piores meses (excluindo a linha TOTAL GERAL)
        comparativo_sem_total = comparativo.drop(index="TOTAL GERAL")
        melhores = comparativo_sem_total.sort_values("VARIAÇÃO (%)", ascending=False).head(3)
        piores = comparativo_sem_total.sort_values("VARIAÇÃO (%)").head(3)

        # ✅ Estilo visual da tabela
        def destacar_total(val):
            return ["font-weight: bold; background-color: #f0f0f0" if val.name == "TOTAL GERAL" else "" for _ in val]

        def destacar_variacao(val):
            if isinstance(val, (int, float)):
                if val > 0:
                    return "color: green; font-weight: bold"
                elif val < 0:
                    return "color: red; font-weight: bold"
            return ""

        # ✅ Aplicar estilos e centralizar cabeçalhos
        comparativo_styled = (
        comparativo.style
        .format({
            "2024": "R$ {:,.2f}",
            "2025": "R$ {:,.2f}",
            "TOTAL GERAL": "R$ {:,.2f}",
            "VARIAÇÃO (%)": "{:.2f}%"
        })
        .apply(destacar_total, axis=1)
        .applymap(destacar_variacao, subset=["VARIAÇÃO (%)"])
        .set_table_styles([
            {"selector": "th", "props": [("text-align", "center")]},
            {"selector": "thead th", "props": [("text-align", "center")]}
        ])
        .set_properties(**{"text-align": "center"})
         )

        st.subheader("📊 Comparativo de Vendas por Ano")
        st.dataframe(comparativo_styled)

        # ✅ Função para destacar variação na narrativa
        def formatar_variacao_html(valor):
            cor = "green" if valor > 0 else "red" if valor < 0 else "black"
            return f"<span style='color:{cor}; font-weight:bold'>{valor:.2f}%</span>"

        # Gerar valores formatados com cor
        variacao_total_html = formatar_variacao_html(variacao_total)
        melhor_mes = melhores.index[0]
        melhor_valor_html = formatar_variacao_html(melhores.iloc[0]["VARIAÇÃO (%)"])
        pior_mes = piores.index[0]
        pior_valor_html = formatar_variacao_html(piores.iloc[0]["VARIAÇÃO (%)"])

        # ✅ Narrativa com destaque visual
        narrativa_html = f"""
        <p style='font-size:16px'>
        Em 2025, as vendas apresentaram uma variação total de {variacao_total_html}em relação ano de 2024.<br>
        Os meses com maior crescimento foram: <strong>{', '.join(melhores.index)}</strong> — com destaque para <strong>{melhor_mes}</strong>, que cresceu {melhor_valor_html}<br>
        Já os meses com pior desempenho foram: <strong>{', '.join(piores.index)}</strong> — sendo <strong>{pior_mes}</strong> o mais crítico, com queda de {pior_valor_html}.
        </p>
        """
        st.subheader("🗣️ Narrativa de Desempenho Anual")
        st.markdown(narrativa_html, unsafe_allow_html=True)

