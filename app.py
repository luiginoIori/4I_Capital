import streamlit as st
import pandas as pd
import os
import re
import json
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import xlrd


#  https://github.com/luiginoIori/4I_Capital.git

def is_date_formatted(cell):
    """Verifica se uma célula está formatada como data"""
    if cell.value is None:
        return False
    
    # Verifica se é datetime
    if isinstance(cell.value, datetime):
        return True
    
    # Verifica se é string que parece data no formato ##/##/####
    if isinstance(cell.value, str):
        date_pattern = r'^\d{1,2}/\d{1,2}/\d{4}$'
        return bool(re.match(date_pattern, cell.value))
    
    # Verifica se é número de data do Excel
    if isinstance(cell.value, (int, float)) and cell.number_format:
        # Formatos de data comuns no Excel
        date_formats = ['dd/mm/yyyy', 'd/m/yyyy', 'mm/dd/yyyy', 'm/d/yyyy', 'yyyy-mm-dd']
        return any(fmt in cell.number_format.lower() for fmt in date_formats)
    
    return False

def process_sicred_files(arquivos):
    """Processa todos os arquivos Sicred de 2025"""
    arq_data =[]
    # Lista arquivos Sicred de 2025
    arquivos_sicred_2025 = []
    for arquivo in arquivos:
        if "2025" in arquivo and "Sicred" in arquivo and (arquivo.endswith('.xls') or arquivo.endswith('.xlsx')):
            arquivos_sicred_2025.append(arquivo)
    
      
    for arquivo in arquivos_sicred_2025:
        caminho_arquivo = arquivo
        
        try:
            # Para arquivos .xls antigos, usa xlrd
            if arquivo.endswith('.xls'):
                workbook_xlrd = xlrd.open_workbook(caminho_arquivo)
                sheet_xlrd = workbook_xlrd.sheet_by_index(0)
                
                
                sair = False
                # Processa linha por linha
                for row_num in range(sheet_xlrd.nrows):
                    try:
                        # Verifica se a célula na coluna A (índice 0) contém uma data
                        col1_val = sheet_xlrd.cell_value(row_num, 0) if sheet_xlrd.ncols > 0 else ""
                        
                        # Verifica se é uma data (número de data do Excel ou string de data)
                        is_date = False
                        if isinstance(col1_val, (int, float)) and col1_val > 0:
                            # Pode ser uma data em formato numérico do Excel
                            try:
                                date_val = xlrd.xldate.xldate_as_datetime(col1_val, workbook_xlrd.datemode)
                                is_date = True
                                col1_val = date_val
                            except:
                                pass
                        elif isinstance(col1_val, str):
                            # Verifica se é string que parece data no formato ##/##/####
                            date_pattern = r'^\d{1,2}/\d{1,2}/\d{4}$'
                            is_date = bool(re.match(date_pattern, col1_val))
                        
                        if is_date:
                            # Se encontrou data formatada, coleta dados das colunas 1, 2, 4
                            col2_val = sheet_xlrd.cell_value(row_num, 1) if sheet_xlrd.ncols > 1 else ""
                            col4_val = sheet_xlrd.cell_value(row_num, 3) if sheet_xlrd.ncols > 3 else ""
                            if col2_val == 'Pag. Boletos':
                                sair = True
                                break
                            x=col1_val,col2_val,col4_val
                            arq_data.append(x)
                            
                    except:
                        continue
                    if sair:
                        break
            
            else:  # Para arquivos .xlsx
                # Carrega o arquivo com openpyxl para verificar formatação
                workbook = load_workbook(caminho_arquivo, data_only=False)
                sheet = workbook.active
                
                st.write(f"Processando: {arquivo}")
                
                # Processa linha por linha
                for row_num in range(1, sheet.max_row + 1):
                    cell_a = sheet.cell(row=row_num, column=1)  # Coluna A (1)
                    
                    if is_date_formatted(cell_a):
                        # Se encontrou data formatada, coleta dados das colunas 1, 2, 4
                        col1_val = sheet.cell(row=row_num, column=1).value
                        col2_val = sheet.cell(row=row_num, column=2).value
                        col4_val = sheet.cell(row=row_num, column=4).value
                        
                        if col2_val == 'Pag. Boletos':
                            sair = True
                            break 
                        x=col1_val,col2_val,col4_val
                        arq_data.append(x)
                    elif row_num > 1 and arq_data:
                        # Se não é data e já coletou dados, verifica se deve parar
                        # Continua coletando até não haver mais datas consecutivas
                        pass
                
                workbook.close()
            
        except Exception as e:
            st.error(f"Erro ao processar {arquivo}: {str(e)}")
    
    return arq_data

def process_bradesco_files(arquivos, arq_data):
    """Processa arquivos Bradesco de 2025"""
      
    # Lista arquivos Bradesco de 2025
    arquivos_bradesco_2025 = []

    for arquivo in arquivos:
        
        if "2025" in arquivo and "Bradesco" in arquivo and (arquivo.endswith('.XLS') or arquivo.endswith('.XLSX') or arquivo.endswith('.xls') or arquivo.endswith('.xlsx')):
            arquivos_bradesco_2025.append(arquivo)
    
    for arquivo in arquivos_bradesco_2025:
        caminho_arquivo = arquivo
        
        try:
            # Para arquivos .xls antigos, usa xlrd
            if arquivo.endswith('.xls') or arquivo.endswith('.XLS'):
                workbook_xlrd = xlrd.open_workbook(caminho_arquivo)
                sheet_xlrd = workbook_xlrd.sheet_by_index(0)
                
                # Procura por "SALDO ANTERIOR"
                saldo_anterior_encontrado = False
                linha_inicio = 0
                total = False
                for row_num in range(sheet_xlrd.nrows):
                    for col_num in range(sheet_xlrd.ncols):                        
                        try:
                            cell_value = sheet_xlrd.cell_value(row_num, col_num)
                            if isinstance(cell_value, str) and "SALDO ANTERIOR" in cell_value.upper():
                                saldo_anterior_encontrado = True
                                linha_inicio = row_num + 1  # Próxima linha
                                break
                        except:
                            continue
                    if saldo_anterior_encontrado:
                        break
                
                if saldo_anterior_encontrado:                    
                    
                    # Coleta dados a partir da linha seguinte
                    for row_num in range(linha_inicio, sheet_xlrd.nrows):
                        try:
                            col1_val = sheet_xlrd.cell_value(row_num, 0) if sheet_xlrd.ncols > 0 else ""
                            col2_val = sheet_xlrd.cell_value(row_num, 1) if sheet_xlrd.ncols > 1 else ""
                            col3_val = sheet_xlrd.cell_value(row_num, 3) if sheet_xlrd.ncols > 3 else ""
                            col4_val = sheet_xlrd.cell_value(row_num, 4) if sheet_xlrd.ncols > 4 else ""
                            
                            # Se coluna 3 está vazia, usa coluna 4
                            terceira_coluna = col4_val if (not col3_val or col3_val == "") else col3_val
                            
                            if col1_val == "Total" or col2_val == None:
                                total = True
                                break
                            x=col1_val,col2_val,terceira_coluna                        
                            arq_data.append(x)                           
                    
                        except:
                            continue
                        if total:
                            break
                    
            else:  # Para arquivos .xlsx
                workbook = load_workbook(caminho_arquivo)
                sheet = workbook.active
                
                # Procura por "SALDO ANTERIOR"
                saldo_anterior_encontrado = False
                linha_inicio = 0
                total = False
                for row_num in range(1, sheet.max_row + 1):
                    for col_num in range(1, sheet.max_column + 1):
                        cell_value = sheet.cell(row=row_num, column=col_num).value
                        if cell_value and isinstance(cell_value, str) and "SALDO ANTERIOR" in cell_value.upper():
                            saldo_anterior_encontrado = True
                            linha_inicio = row_num + 1  # Próxima linha
                            break
                    if saldo_anterior_encontrado:
                        break
                
                if saldo_anterior_encontrado:                    
                    
                    # Coleta dados a partir da linha seguinte
                    for row_num in range(linha_inicio, sheet.max_row + 1):
                        col1_val = sheet.cell(row=row_num, column=1).value
                        col2_val = sheet.cell(row=row_num, column=2).value
                        col3_val = sheet.cell(row=row_num, column=4).value
                        col4_val = sheet.cell(row=row_num, column=5).value
                       
                        # Se coluna 3 está vazia, usa coluna 4
                        terceira_coluna = col4_val if (not col3_val or col3_val == "") else col3_val
                        if col1_val == "Total" or col2_val == None:
                            total = True
                        x=col1_val,col2_val,terceira_coluna   
                        if total:
                            break                     
                        arq_data.append(x)
                        
                workbook.close()
        
        except Exception as e:
            st.error(f"Erro ao processar {arquivo}: {str(e)}")
        
    return arq_data



def descricao(df_bradesco):
       
    # Processar e alterar os itens diretamente na lista df_bradesco
    df_bradesco_atualizado = []
    
    for i in df_bradesco:
        if len(i) > 1:  # Verifica se a tupla tem pelo menos 2 elementos
            # Pega o item 1 e converte para string
            item_original = str(i[1]) if i[1] is not None else ""
            
            # Divide o texto por espaços
            partes = item_original.split()
            
            # Filtra as partes, removendo apenas as que contêm SOMENTE números
            partes_filtradas = []
            # Lista de palavras a serem removidas
            palavras_remover = ['PAGAMENTO', 'PIX', 'ELETRON', 'COBRANCA', 'PAGTO', 'ELETRÔNICO', 'RECEBIMENTO', 
                                'TRANSF', 'TED', 'TRANSFERÊNCIA','BOLETO DE LIQUIDAÇÃO', 'LIQUIDAÇÃO',
                                'BOLETO', 'LIQUIDACAO','SICREDI ']
            
            for parte in partes:
                # Remove caracteres especiais da parte para análise
                parte_limpa = re.sub(r'[^A-Za-zÀ-ÿ0-9]', '', parte)
                parte_upper = parte_limpa.upper()
                
                # Verifica se a parte não é uma das palavras a serem removidas
                if parte_upper not in palavras_remover:
                    # Se a parte limpa não é apenas números OU se contém letras, mantém
                    if not parte_limpa.isdigit() or any(c.isalpha() for c in parte_limpa):
                        # Remove caracteres especiais mas mantém a parte
                        parte_sem_especiais = re.sub(r'[^A-Za-zÀ-ÿ0-9]', '', parte)
                        if parte_sem_especiais:  # Só adiciona se não ficou vazia
                            partes_filtradas.append(parte_sem_especiais)
            
            # Reconstrói o texto com espaços entre as posições mantidas
            item_limpo = ' '.join(partes_filtradas)
            
            # Converte para maiúsculas
            item_limpo = item_limpo.upper()
            
            # Cria nova tupla com o item 1 tratado
            if len(i) >= 3:
                nova_tupla = (i[0], item_limpo if item_limpo else i[1], i[2])
            else:
                nova_tupla = (i[0], item_limpo if item_limpo else i[1])
            
            # Aplicar padrões de descrição
            if len(nova_tupla) > 1 and nova_tupla[1]:
                descricao_padronizada = str(nova_tupla[1])
                
                # BARRIER para BARRIER TERCEIRIZACAO
                if "BARRIER" in descricao_padronizada.upper():
                    descricao_padronizada = "BARRIER TERCEIRIZACAO"
                # AURIGA para AURIGA FUNDO DE INVESTIMENTO
                elif "AURIGA" in descricao_padronizada.upper():
                    descricao_padronizada = "AURIGA FUNDO DE INVESTIMENTO"
                # BB CLAIM para BB CLAIM FUNDO DE INVESTIMENTO
                elif "BB CLAIM" in descricao_padronizada.upper():
                    descricao_padronizada = "BB CLAIM FUNDO DE INVESTIMENTO"
                # EVOLUCAO para EVOLUCAO AUDITORES
                elif "EVOLUCAO" in descricao_padronizada.upper():
                    descricao_padronizada = "EVOLUCAO AUDITORES"
                # EXCELSIOR para EXCELSIOR FUNDO DE INVESTIMENTO
                elif "EXCELSIOR" in descricao_padronizada.upper():
                    descricao_padronizada = "EXCELSIOR FUNDO DE INVESTIMENTO"
                # JL MORAIS para JL MORAIS SOLUCOES
                elif "MORAIS" in descricao_padronizada.upper():
                    descricao_padronizada = "JL MORAIS SOLUCOES"
                # SBC OPORTUNIDADE para SBC OPORTUNIDADE FUNDO DE
                elif "SBC OPORTUNIDADE" in descricao_padronizada.upper():
                    descricao_padronizada = "SBC OPORTUNIDADE FUNDO DE"
                # TEDTRANSF ELET DISPON REMETNDMP para NDMP I FIDC
                elif "TEDTRANSF ELET DISPON REMETNDMP" in descricao_padronizada.upper():
                    descricao_padronizada = "NDMP I FIDC"
                # TRANSFERÊNCIA DE 3I HOLDING LTDA para 3I HOLDING LTDA
                elif "3I" in descricao_padronizada.upper():
                    descricao_padronizada = "3I HOLDING LTDA"
                # TRANSFERÊNCIA DES 4I CAPITAL LTDA para 4I CAPITAL LTDA
                elif "4I CAPITAL" in descricao_padronizada.upper():
                    descricao_padronizada = "4I CAPITAL LTDA"
                # TRANSFERÊNCIA DE IGOR JEFFERSON LIMA C para IGOR JEFFERSON LIMA C
                elif "IGOR JEFFERSON" in descricao_padronizada.upper():
                    descricao_padronizada = "IGOR JEFFERSON LIMA C"
                # PROCESSO ANBIMA para ANBIMA ASSOC BR
                elif "ANBIMA" in descricao_padronizada.upper():
                    descricao_padronizada = "ANBIMA ASSOC BR"
                # NIO DIGITAL para NDMP I FIDC
                elif "NIO DIGITAL" in descricao_padronizada.upper():
                    descricao_padronizada = "NDMP I FIDC"
                # VANEY para VANEY IORI
                elif "VANEY" in descricao_padronizada.upper():
                    descricao_padronizada = "VANEY IORI"
                # LOCALIZA para LOCALIZA FLEET S A
                elif "FLEET" in descricao_padronizada.upper():
                    descricao_padronizada = "LOCALIZA FLEET S A"
                # PREFEITURA para PREFEITURA MUNI
                elif "PREFEITURA" in descricao_padronizada.upper():
                    descricao_padronizada = "PREFEITURA MUNICIPAL"
                # TRIAGEM para TRIAGEM CONSULTORIA
                elif "TRIAGEM" in descricao_padronizada.upper():
                    descricao_padronizada = "TRIAGEM CONSULTORIA"
                # V IORI para V IORI ADVISORY
                elif "V IORI" in descricao_padronizada.upper():
                    descricao_padronizada = "V IORI ADVISORY"
                # RENTABINVEST para INVEST FACEL
                elif "RENTABINVEST" in descricao_padronizada.upper():
                    descricao_padronizada = "INVEST FACIL"
                # LIF DESENVO para LIF DESENVOLVIMENTO
                elif "LIF DESENVO" in descricao_padronizada.upper():
                    descricao_padronizada = "LIF DESENVOLVIMENTO"
                # LUIGINO para LUIGINO IORI FILHO
                elif "LUIGINO" in descricao_padronizada.upper():
                    descricao_padronizada = "LUIGINO IORI FILHO"
                # TARIFA para TARIFA BANCARIA
                elif "TARIFA" in descricao_padronizada.upper():
                    descricao_padronizada = "TARIFA BANCARIA"
                # OPERACAO CAPITAL GIRO para OPERACAO CAPITAL GIRO
                elif "CAPITAL GIRO" in descricao_padronizada.upper():
                    descricao_padronizada = "OPERACAO CAPITAL GIRO"
                
                # Atualizar tupla com descrição padronizada
                if descricao_padronizada != str(nova_tupla[1]):
                    if len(nova_tupla) >= 3:
                        nova_tupla = (nova_tupla[0], descricao_padronizada, nova_tupla[2])
                    else:
                        nova_tupla = (nova_tupla[0], descricao_padronizada)
            
            df_bradesco_atualizado.append(nova_tupla)
        else:
            # Se a tupla não tem item 1, mantém como está
            df_bradesco_atualizado.append(i)
        
    return df_bradesco_atualizado


def carregar_classificacoes():
    """Carrega classificações existentes do arquivo JSON"""
    arquivo_classificacoes = "classificacoes_descricoes.json"
    if os.path.exists(arquivo_classificacoes):
        try:
            with open(arquivo_classificacoes, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def salvar_classificacoes(classificacoes):
    """Salva classificações no arquivo JSON"""
    arquivo_classificacoes = "classificacoes_descricoes.json"
    try:
        with open(arquivo_classificacoes, 'w', encoding='utf-8') as f:
            json.dump(classificacoes, f, ensure_ascii=False, indent=2)
        return True
    except:
        return False


def obter_descricoes_unicas(dados_completos):
    """Obtém lista de descrições únicas dos dados"""
    descricoes = set()
    for registro in dados_completos:
        if len(registro) >= 2 and registro[1]:
            descricoes.add(str(registro[1]).strip())
    return sorted(list(descricoes))


def formulario_classificacao(dados_completos):
    """Cria formulário para classificação das descrições"""
    st.markdown("---")
    st.subheader("📝 Classificação de Descrições")
    
    # Opções de classificação disponíveis
    opcoes_classificacao = [
        "",
        "RECEITAS",
        "EMPRÉSTIMOS",
        "REEMBOLSO", 
        "CONTA CORRENTE",
        "APLICAÇÃO FINANCEIRA",
        "TRANSFERENCIA ENTRE CONTAS",
        "IMPOSTOS",
        "FOLHA CLT",
        "FOLHA PJ",
        "ENCARGOS",
        "ADMINISTRATIVA",
        "ASSESSORIA JURIDICA",
        "ASSESSORIA CONTABIL",
        "DESPESAS FINANCEIRAS",
        "DESPESAS COMERCIAIS",
        "SOFTWARE",
        "PMT EMPRESTIMOS",
        "INVESTIMENTOS",
        "DESPESAS IMÓVEL",        
        "ADIANTAMENTO A FORNECEDORES"
        
    ]
    
    # Carregar classificações existentes
    classificacoes_existentes = carregar_classificacoes()
    
    # Obter descrições únicas
    descricoes_unicas = obter_descricoes_unicas(dados_completos)
    
    # Filtrar descrições não classificadas
    descricoes_nao_classificadas = [desc for desc in descricoes_unicas 
                                    if desc not in classificacoes_existentes]
    
    # Adicionar seletor para editar classificações existentes
    st.subheader("✏️ Editar Classificações Existentes")
    
    if classificacoes_existentes:
        # Selectbox para escolher descrição para editar
        descricoes_para_editar = ["Selecione uma descrição para editar..."] + sorted(list(classificacoes_existentes.keys()))
        descricao_selecionada = st.selectbox(
            "Escolha uma descrição para editar:",
            descricoes_para_editar,
            key="edit_selector"
        )
        
        if descricao_selecionada != "Selecione uma descrição para editar...":
            classificacao_atual = classificacoes_existentes[descricao_selecionada]
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.write(f"**Descrição:** {descricao_selecionada}")
                st.write(f"**Classificação atual:** {classificacao_atual}")
                
                # Dropdown para nova classificação
                try:
                    index_atual = opcoes_classificacao.index(classificacao_atual)
                except ValueError:
                    index_atual = 0
                
                nova_classificacao = st.selectbox(
                    "Nova classificação:",
                    opcoes_classificacao,
                    index=index_atual,
                    key="new_classification"
                )
            
            with col2:
                st.write("")
                st.write("")
                
                if st.button("💾 Atualizar Classificação", key="update_btn"):
                    if nova_classificacao != classificacao_atual:
                        classificacoes_existentes[descricao_selecionada] = nova_classificacao
                        if salvar_classificacoes(classificacoes_existentes):
                            st.success(f"✅ Classificação atualizada para: **{nova_classificacao}**")
                            st.rerun()
                        else:
                            st.error("❌ Erro ao salvar a atualização!")
                    else:
                        st.info("ℹ️ Nenhuma alteração detectada.")
                
                if st.button("🗑️ Excluir Classificação", key="delete_btn"):
                    del classificacoes_existentes[descricao_selecionada]
                    if salvar_classificacoes(classificacoes_existentes):
                        st.success("✅ Classificação excluída!")
                        st.rerun()
                    else:
                        st.error("❌ Erro ao excluir!")
    
    st.markdown("---")
    
    if not descricoes_nao_classificadas:
        st.success("✅ Todas as descrições já estão classificadas!")
        
        # Mostrar classificações existentes
        if st.checkbox("Mostrar todas as classificações cadastradas"):
            st.write("**Classificações cadastradas:**")
            for desc, classif in sorted(classificacoes_existentes.items()):
                st.write(f"• {desc} → **{classif}**")
        return
    
    st.subheader("➕ Classificar Novas Descrições")
    st.write(f"**{len(descricoes_nao_classificadas)}** descrições precisam ser classificadas:")
    
    # Formulário para classificar
    with st.form("classificacao_form"):
        classificacoes_novas = {}
        
        # Dividir em colunas para melhor layout
        num_cols = 2
        cols = st.columns(num_cols)
        
        for i, descricao in enumerate(descricoes_nao_classificadas[:10]):  # Limitar a 10 por vez
            col = cols[i % num_cols]
            
            with col:
                st.write(f"**Descrição:** {descricao}")
                classificacao = st.selectbox(
                    "Classificação:",
                    ["Selecione..."] + opcoes_classificacao,
                    key=f"class_{i}",
                    index=0
                )
                
                if classificacao != "Selecione...":
                    classificacoes_novas[descricao] = classificacao
                
                st.write("---")
        
        # Botões do formulário
        col1, col2 = st.columns([1, 1])
        
        with col1:
            submitted = st.form_submit_button("💾 Salvar Classificações")
        
        with col2:
            if len(descricoes_nao_classificadas) > 10:
                st.write(f"Restam {len(descricoes_nao_classificadas) - 10} descrições")
    
    # Processar envio do formulário
    if submitted and classificacoes_novas:
        # Mesclar com classificações existentes
        classificacoes_existentes.update(classificacoes_novas)
        
        # Salvar no arquivo
        if salvar_classificacoes(classificacoes_existentes):
            st.success(f"✅ {len(classificacoes_novas)} classificações salvas com sucesso!")
            st.rerun()
        else:
            st.error("❌ Erro ao salvar classificações!")


def aplicar_classificacoes(dados_completos):
    """Aplica classificações aos dados e retorna dados classificados"""
    classificacoes = carregar_classificacoes()
    dados_classificados = []
    
    for registro in dados_completos:
        if len(registro) >= 3:
            data, descricao, valor = registro[0], registro[1], registro[2]
            classificacao = classificacoes.get(str(descricao).strip(), "NÃO CLASSIFICADO")
            # Adiciona classificação como quarta coluna
            dados_classificados.append((data, descricao, valor, classificacao))
        else:
            dados_classificados.append(registro + ("NÃO CLASSIFICADO",))
    
    return dados_classificados


def criar_tabela_por_classificacao(dados_classificados):
    """Cria tabela resumo por classificação"""
    if not dados_classificados:
        return ""
    
    # Organizar dados por classificação
    resumo_classificacao = {}
    
    for registro in dados_classificados:
        if len(registro) >= 4:
            data, descricao, valor, classificacao = registro[0], registro[1], registro[2], registro[3]
            
            # Converter valor para float
            if isinstance(valor, str):
                valor_limpo = re.sub(r'[^\d,.\-]', '', str(valor))
                valor_limpo = valor_limpo.replace(',', '.')
                try:
                    valor_float = float(valor_limpo)
                except:
                    valor_float = 0.0
            else:
                valor_float = float(valor) if valor else 0.0
            
            if classificacao not in resumo_classificacao:
                resumo_classificacao[classificacao] = 0.0
            
            resumo_classificacao[classificacao] += valor_float
    
    # Criar HTML da tabela de classificação
    html = """
    <div style="margin-top: 20px;">
    <h3>💼 Resumo por Classificação</h3>
    <table style="border-collapse: collapse; width: 100%; max-width: 800px;">
    <thead>
        <tr style="background-color: #f0f0f0;">
            <th style="border: 1px solid #ddd; padding: 12px; text-align: left;">Classificação</th>
            <th style="border: 1px solid #ddd; padding: 12px; text-align: right;">Total (R$)</th>
        </tr>
    </thead>
    <tbody>
    """
    
    # Ordenar por classificação
    for classificacao in sorted(resumo_classificacao.keys()):
        total = resumo_classificacao[classificacao]
        cor_texto = "color: #dc3545;" if total < 0 else "color: #28a745;"
        
        html += f"""
        <tr>
            <td style="border: 1px solid #ddd; padding: 8px;">{classificacao}</td>
            <td style="border: 1px solid #ddd; padding: 8px; text-align: right; {cor_texto}">
                R$ {total:,.2f}
            </td>
        </tr>
        """
    
    html += """
    </tbody>
    </table>
    </div>
    """
    
    return html


def criar_tabela_mensal(dados_completos):
    """Cria tabela HTML organizada por descrição e meses"""
    
    # Ordem das classificações para ordenação
    ordem_classificacoes = [
        "RECEITAS",
        "EMPRÉSTIMOS",
        "REEMBOLSO", 
        "CONTA CORRENTE",
        "APLICAÇÃO FINANCEIRA",
        "TRANSFERENCIA ENTRE CONTAS",
        "DESPESAS",
        "IMPOSTOS",
        "FOLHA CLT",
        "FOLHA PJ",
        "ENCARGOS",
        "ADMINISTRATIVA",
        "ASSESSORIA JURIDICA",
        "ASSESSORIA CONTABIL",
        "DESPESAS FINANCEIRAS",
        "DESPESAS COMERCIAIS",
        "SOFTWARE",
        "PMT EMPRESTIMOS",
        "INVESTIMENTOS",
        "DESPESAS IMÓVEL",        
        "ADIANTAMENTO A FORNECEDORES",
        "NÃO CLASSIFICADO"
    ]
    
    # Definir subcategorias de DESPESAS
    subcategorias_despesas = [
        "IMPOSTOS",
        "FOLHA CLT",
        "FOLHA PJ",
        "ENCARGOS",
        "ADMINISTRATIVA",
        "ASSESSORIA JURIDICA",
        "ASSESSORIA CONTABIL",
        "DESPESAS FINANCEIRAS",
        "DESPESAS COMERCIAIS",
        "SOFTWARE",
        "PMT EMPRESTIMOS",
        "INVESTIMENTOS",
        "DESPESAS IMÓVEL",
        "ADIANTAMENTO A FORNECEDORES"
    ]
    
    # Carregar classificações para ordenação
    classificacoes = carregar_classificacoes()
    
    # Dicionário para armazenar os dados organizados
    tabela_dados = {}
    
    # Processar cada registro
    for registro in dados_completos:
        if len(registro) >= 3:
            data, descricao, valor = registro[0], registro[1], registro[2]
            
            # Extrair mês da data
            try:
                if isinstance(data, datetime):
                    mes = data.month
                elif isinstance(data, str):
                    # Tentar diferentes formatos de data
                    for formato in ['%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y']:
                        try:
                            data_obj = datetime.strptime(data, formato)
                            mes = data_obj.month
                            break
                        except:
                            continue
                    else:
                        mes = 1  # Default para janeiro se não conseguir parsear
                else:
                    mes = 1  # Default
                
                # Converter valor para float
                if isinstance(valor, str):                
                    # Remove caracteres não numéricos exceto vírgula, ponto e sinal negativo
                    valor_limpo = re.sub(r'[^\d,.\-]', '', str(valor))
                    valor_limpo = valor_limpo.replace(',', '.')
                    
                    try:
                        valor_float = float(valor_limpo)
                        
                    except:
                        valor_float = 0.0
                        
                else:
                    valor_float = float(valor) if valor else 0.0                
                if descricao not in tabela_dados:                    
                    tabela_dados[descricao] = {i: 0.0 for i in range(1, 13)}  # Meses 1-12
                
                # Somar valor ao mês correspondente                
                tabela_dados[descricao][mes] += valor_float                
                
            except Exception as e:
                print(f"Erro processando registro: {e}")
                continue
    
    # Criar HTML da tabela - versão simplificada
    meses_nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                   'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    
    html = """
    <div style="max-height: 600px; overflow-y: auto;">
    <table border="1" style="width:100%; border-collapse: collapse;">
    <tr style="background-color: #f0f0f0;">
        <th style="text-align: left; padding: 8px; background-color: #f0f0f0;">Descrição</th>
        <th style="text-align: center; padding: 8px; background-color: #f0f0f0;">Total</th>
    """
    
    # Adicionar cabeçalhos dos meses
    for mes_nome in meses_nomes:
        html += f'<th style="text-align: center; padding: 8px;">{mes_nome}</th>'
    
    html += """
    </tr>
    """
    
    # Verificar se há dados para processar
    if not tabela_dados:
        return "<p>Nenhum dado encontrado para gerar a tabela mensal.</p>"
    
    # Definir classificações para linha de total especial
    classificacoes_receitas = [
        "RECEITAS",
        "EMPRÉSTIMOS",
        "REEMBOLSO", 
        "CONTA CORRENTE",
        "APLICAÇÃO FINANCEIRA",
        "TRANSFERENCIA ENTRE CONTAS"
    ]
    
    # Calcular totais das classificações de receitas
    def calcular_totais_receitas():
        totais_mes_receitas = {i: 0.0 for i in range(1, 13)}  # Meses 1-12
        total_geral_receitas = 0.0
        
        for desc in tabela_dados.keys():
            classificacao_desc = classificacoes.get(str(desc).strip(), "NÃO CLASSIFICADO")
            if classificacao_desc in classificacoes_receitas:
                dados_mes = tabela_dados[desc]
                for mes in range(1, 13):
                    totais_mes_receitas[mes] += dados_mes[mes]
                total_geral_receitas += sum(dados_mes.values())
        
        return totais_mes_receitas, total_geral_receitas
    
    # Calcular totais das despesas para o saldo bancário
    def calcular_totais_despesas():
        totais_mes_despesas = {i: 0.0 for i in range(1, 13)}  # Meses 1-12
        total_geral_despesas = 0.0
        
        for desc in tabela_dados.keys():
            classificacao_desc = classificacoes.get(str(desc).strip(), "NÃO CLASSIFICADO")
            if classificacao_desc in subcategorias_despesas:
                dados_mes = tabela_dados[desc]
                for mes in range(1, 13):
                    totais_mes_despesas[mes] += dados_mes[mes]
                total_geral_despesas += sum(dados_mes.values())
        
        return totais_mes_despesas, total_geral_despesas
    
    # Calcular totais das receitas e despesas
    totais_mes_receitas, total_geral_receitas = calcular_totais_receitas()
    totais_mes_despesas, total_geral_despesas = calcular_totais_despesas()
    
    # Calcular saldo bancário cumulativo
    saldo_inicial = 272801.75
    saldo_acumulado = saldo_inicial
    saldos_mensais_cumulativos = {}
    
    # Calcular saldo acumulado mês a mês
    for mes in range(1, 13):
        movimento_mes = totais_mes_receitas[mes] + totais_mes_despesas[mes]  # receitas + despesas (despesas são negativas)
        saldo_acumulado += movimento_mes
        saldos_mensais_cumulativos[mes] = saldo_acumulado
    
    # Adicionar linha de saldo bancário (sem coluna total)
    html += f'<tr style="background-color: rgba(255, 215, 0, 0.4); font-weight: bold; border: 3px solid #FFD700;">'
    html += f'<td style="padding: 8px; text-align: center; font-size: 20px; font-weight: bold; color: #4169E1;">💰 SALDO BANCÁRIO</td>'
    html += f'<td style="padding: 8px; text-align: center; font-size: 16px; font-weight: bold; color: #666;">Saldo Inicial: {int(saldo_inicial):,}</td>'
    
    # Adicionar saldo acumulado de cada mês
    for mes in range(1, 13):
        saldo_mes_acumulado = saldos_mensais_cumulativos[mes]
        cor_saldo_mes = '#0066CC' if saldo_mes_acumulado >= 0 else '#DC143C'  # Azul para positivos, vermelho para negativos
        html += f'<td style="padding: 8px; color: {cor_saldo_mes}; text-align: center; font-weight: bold; font-size: 18px;">{int(saldo_mes_acumulado):,}</td>'
    
    html += '</tr>'
    
    # Adicionar linha de divisão estreita
    html += f'<tr style="height: 3px; border: none;"><td colspan="14" style="background-color: #666; height: 3px; border: none; padding: 0;"></td></tr>'
    
    # Adicionar linha RECEITA / DESPESAS (soma de receitas + despesas)
    total_receita_despesas = total_geral_receitas + total_geral_despesas
    html += f'<tr style="background-color: rgba(255, 165, 0, 0.4); font-weight: bold; border: 2px solid #FFA500;">'
    html += f'<td style="padding: 8px; text-align: center; font-size: 18px; font-weight: bold; color: #FF4500;">💰 RECEITA / DESPESAS</td>'
    
    # Calcular cor para o total (positivo = verde, negativo = vermelho)
    cor_total_receita_despesas = '#228B22' if total_receita_despesas >= 0 else '#DC143C'
    html += f'<td style="padding: 8px; color: {cor_total_receita_despesas}; text-align: center; font-size: 18px; font-weight: bold;">{int(total_receita_despesas):,}</td>'
    
    # Adicionar valores mensais (receitas + despesas por mês)
    for mes in range(1, 13):
        valor_mes_total = totais_mes_receitas[mes] + totais_mes_despesas[mes]
        if valor_mes_total != 0:
            cor_valor = '#228B22' if valor_mes_total >= 0 else '#DC143C'
            html += f'<td style="padding: 8px; color: {cor_valor}; text-align: center; font-weight: bold; font-size: 18px;">{int(valor_mes_total):,}</td>'
        else:
            html += '<td style="padding: 8px; text-align: center; font-weight: bold; font-size: 18px; color: #FF4500;">-</td>'
    
    html += '</tr>'
    
    # Adicionar linha de divisão estreita
    html += f'<tr style="height: 3px; border: none;"><td colspan="14" style="background-color: #666; height: 3px; border: none; padding: 0;"></td></tr>'
    
    # Adicionar linha de total das receitas
    html += f'<tr style="background-color: rgba(144, 238, 144, 0.6); font-weight: bold; border: 2px solid #90EE90;">'
    html += f'<td style="padding: 8px; text-align: center; font-size: 18px; font-weight: bold; color: #000080;">💰 RECEITAS/APLICAÇÕES/EMPRESTIMOS</td>'
    html += f'<td style="padding: 8px; color: #000080; text-align: center; font-size: 18px; font-weight: bold;">{int(total_geral_receitas):,}</td>'
    
    # Adicionar totais de cada mês para receitas
    for mes in range(1, 13):
        valor_mes = totais_mes_receitas[mes]
        if valor_mes != 0:
            html += f'<td style="padding: 8px; color: #000080; text-align: center; font-weight: bold; font-size: 18px;">{int(valor_mes):,}</td>'
        else:
            html += '<td style="padding: 8px; text-align: center; font-weight: bold; font-size: 18px; color: #000080;">-</td>'
    
    html += '</tr>'
    
    # Adicionar linha de divisão mais grossa entre RECEITAS/APLICAÇÕES e dados detalhados
    html += f'<tr style="height: 8px; border: none;"><td colspan="14" style="background-color: #00CED1; height: 8px; border: none; padding: 0; border-top: 3px solid #008B8B; border-bottom: 2px solid #008B8B;"></td></tr>'
    
    # Função para obter índice da classificação para ordenação
    def obter_indice_classificacao(descricao):
        classificacao = classificacoes.get(str(descricao).strip(), "NÃO CLASSIFICADO")
        try:
            return ordem_classificacoes.index(classificacao)
        except ValueError:
            return len(ordem_classificacoes)  # Colocar no final se não encontrar
    
    # Ordenar descrições por classificação e depois alfabeticamente dentro da mesma classificação
    descricoes_ordenadas = sorted(tabela_dados.keys(), 
                                 key=lambda desc: (obter_indice_classificacao(desc), desc))
    
    # Calcular totais por classificação
    def calcular_totais_classificacao(classificacao):
        totais_mes = {i: 0.0 for i in range(1, 13)}  # Meses 1-12
        total_geral = 0.0
        
        # Se é DESPESAS, somar todas as subcategorias
        if classificacao == "DESPESAS":
            for desc in descricoes_ordenadas:
                classificacao_desc = classificacoes.get(str(desc).strip(), "NÃO CLASSIFICADO")
                if classificacao_desc in subcategorias_despesas:
                    dados_mes = tabela_dados[desc]
                    for mes in range(1, 13):
                        totais_mes[mes] += dados_mes[mes]
                    total_geral += sum(dados_mes.values())
        else:
            # Lógica normal para outras classificações
            for desc in descricoes_ordenadas:
                classificacao_desc = classificacoes.get(str(desc).strip(), "NÃO CLASSIFICADO")
                if classificacao_desc == classificacao:
                    dados_mes = tabela_dados[desc]
                    for mes in range(1, 13):
                        totais_mes[mes] += dados_mes[mes]
                    total_geral += sum(dados_mes.values())
        
        return totais_mes, total_geral

    # Adicionar dados de cada descrição ordenada com separadores por classificação
    classificacao_anterior = None
    despesas_ja_adicionada = False
    
    for descricao in descricoes_ordenadas:
        # Obter classificação atual
        classificacao_atual = classificacoes.get(str(descricao).strip(), "NÃO CLASSIFICADO")
        
        # Se é uma subcategoria de DESPESAS e ainda não foi adicionada a linha DESPESAS
        if classificacao_atual in subcategorias_despesas and not despesas_ja_adicionada:
            # Adicionar linha DESPESAS primeiro
            totais_mes_despesas, total_despesas = calcular_totais_classificacao("DESPESAS")
            
            cor_total_despesas = 'red' if total_despesas < 0 else 'green'
            html += f'<tr style="background-color: rgba(64, 224, 208, 0.3); font-weight: bold; border: 2px solid rgba(64, 224, 208, 0.8);">'
            html += f'<td style="padding: 8px; text-align: center; font-size: 18px; font-weight: bold;">💰 DESPESAS MENSAIS</td>'
            html += f'<td style="padding: 8px; color: {cor_total_despesas}; text-align: center; font-size: 18px; font-weight: bold;">{int(total_despesas):,}</td>'
            
            # Adicionar totais de cada mês para DESPESAS
            for mes in range(1, 13):
                valor_mes = totais_mes_despesas[mes]
                if valor_mes != 0:
                    cor_valor = 'red' if valor_mes < 0 else 'green'
                    html += f'<td style="padding: 8px; color: {cor_valor}; text-align: center; font-weight: bold; font-size: 18px;">{int(valor_mes):,}</td>'
                else:
                    html += '<td style="padding: 8px; text-align: center; font-weight: bold; font-size: 18px;">-</td>'
            
            html += '</tr>'
            
            # Adicionar linha de divisão estreita após DESPESAS (TOTAL)
            html += f'<tr style="height: 3px; border: none;"><td colspan="14" style="background-color: #666; height: 3px; border: none; padding: 0;"></td></tr>'
            
            despesas_ja_adicionada = True
        
        # Se mudou a classificação, adicionar linha separadora com totais
        if classificacao_atual != classificacao_anterior:
            # Calcular totais para esta classificação
            totais_mes_classificacao, total_classificacao = calcular_totais_classificacao(classificacao_atual)
            
            # Linha da classificação com totais
            cor_total_classif = 'red' if total_classificacao < 0 else 'green'
            html += f'<tr style="background-color: #e8f4fd; font-weight: bold;">'
            html += f'<td style="padding: 8px; text-align: center;">📋 {classificacao_atual}</td>'
            html += f'<td style="padding: 8px; color: {cor_total_classif}; text-align: center;">{int(total_classificacao):,}</td>'
            
            # Adicionar totais de cada mês para a classificação
            for mes in range(1, 13):
                valor_mes = totais_mes_classificacao[mes]
                if valor_mes != 0:
                    cor_valor = 'red' if valor_mes < 0 else 'green'
                    html += f'<td style="padding: 8px; color: {cor_valor}; text-align: center; font-weight: bold;">{int(valor_mes):,}</td>'
                else:
                    html += '<td style="padding: 8px; text-align: center; font-weight: bold;">-</td>'
            
            html += '</tr>'
            classificacao_anterior = classificacao_atual
        
        dados_mes = tabela_dados[descricao]
        total_descricao = sum(dados_mes.values())
        
        cor_total = 'red' if total_descricao < 0 else 'green'
        html += f'<tr style="height: 30px;"><td style="padding: 4px 8px;">{descricao}</td>'
        html += f'<td style="padding: 4px 8px; color: {cor_total}; text-align: center;">{int(total_descricao):,}</td>'
        
        # Adicionar valores de cada mês
        for mes in range(1, 13):
            valor = dados_mes[mes]
            if valor != 0:
                cor_valor = 'red' if valor < 0 else 'black'
                html += f'<td style="padding: 4px 8px; color: {cor_valor}; text-align: center;">{int(valor):,}</td>'
            else:
                html += '<td style="padding: 4px 8px; text-align: center;">-</td>'
        
        html += '</tr>'
    
    html += """
    </table>
    """    
    return html


def main():
    # Configurar layout da página para usar toda a largura
    st.set_page_config(
        page_title="Processador de Extratos Excel - 4I Capital Ltda.",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # Título centralizado, grande e em azul turquesa
    st.markdown("""
    <h1 style='text-align: center; color: #40E0D0; font-size: 3rem; font-weight: bold; margin-bottom: 1rem;'>
        Processador de Extratos Excel - Sicred e Bradesco 2025 - 4I Capital Ltda.
    </h1>
    """, unsafe_allow_html=True)
    
    arquivos = arquivos_disponiveis()    
    # Botão para processar arquivos
    dados_sicred = process_sicred_files(arquivos)
    arquivos = arquivos_disponiveis()
    dados_bradesco = process_bradesco_files(arquivos,dados_sicred)
    dados_completos = descricao(dados_bradesco)
    
    # Formulário de classificação das descrições
    if dados_completos:
        formulario_classificacao(dados_completos)    
    
    # Criar e exibir tabela HTML mensal
    if dados_completos:
        st.markdown("---")
        st.subheader("📊 Tabela Mensal por Descrição")
        st.markdown("")
        tabela_html = criar_tabela_mensal(dados_completos)
        st.markdown(tabela_html, unsafe_allow_html=True)
  
def arquivos_disponiveis():
    # Informações sobre os arquivos na pasta    
    arquivos_dir = "ArquivosExtratos"   
    if os.path.exists(arquivos_dir):
        arquivos = os.listdir(arquivos_dir)        
        arquivos_2025 = [arq for arq in arquivos if "2025" in arq]                 
        if arquivos_2025:            
            for arq in arquivos_2025:
                tipo = "Sicred" if "Sicred" in arq else "Bradesco" if "Bradesco" in arq else "Outro"
                
        else:
            st.write("Nenhum arquivo de 2025 encontrado.")
    else:
        st.error(f"Pasta {arquivos_dir} não encontrada!")
    import pandas as pd

    path = arquivos_dir    
    arquivos_dir = arquivos_2025    
    df = pd.DataFrame(arquivos_dir)
    st.dataframe(df)
    return arquivos_dir

if __name__ == "__main__":

    main()










