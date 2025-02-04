import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import pdfplumber
import io

def extract_table_from_pdf(pdf_file):
    table_data = []
    extracting = False
    pdf_bytes = pdf_file.read()  # Lendo o PDF uma única vez na memória
    pdf_stream = io.BytesIO(pdf_bytes)  # Criando um fluxo de bytes reutilizável
    
    # Primeiro, tentar com pdfplumber para uma extração mais robusta
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            tables = page.extract_table()
            if tables:
                for row in tables:
                    if row and row[0] and isinstance(row[0], str) and row[0].strip().isdigit():  # Verificar se a linha é válida
                        table_data.append(row)
    
    # Se a extração com pdfplumber falhar, tentar com PyMuPDF
    if not table_data:
        pdf_stream.seek(0)  # Resetando o fluxo para permitir nova leitura
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        current_item = []
        for page in doc:
            text = page.get_text("text")
            lines = text.split("\n")
            
            for line in lines:
                line = line.strip()
                if "LISTA DE MATERIAIS" in line and "MACROMEDIDOR" in line:
                    extracting = True  # Começa a extrair dados a partir desse ponto
                    continue
                
                if extracting:
                    parts = line.split()
                    if len(parts) >= 3 and parts[0].isdigit():  # Garante que a linha começa com um número (ITEM)
                        if current_item:
                            table_data.append(current_item)  # Salva a linha anterior antes de iniciar a nova
                        current_item = [parts[0]]  # Inicia nova linha com o item
                        descricao_part = []
                        
                        # Processar os campos da tabela
                        for i, part in enumerate(parts[1:], start=1):
                            if part.isalpha() or "-" in part:  # Parte do nome ou material
                                descricao_part.append(part)
                            elif part.isdigit() or "x" in part:  # Número ou dimensão
                                if len(current_item) < 8:
                                    current_item.append(part)
                                else:
                                    descricao_part.append(part)
                        descricao = " ".join(descricao_part)
                        current_item.insert(1, descricao)
                    elif len(parts) == 1 and extracting and current_item:
                        # Caso a linha contenha um único elemento, pode ser parte da descrição anterior
                        current_item[1] += " " + parts[0]
                    elif "NOTAS" in line:
                        extracting = False
                        if current_item:
                            table_data.append(current_item)  # Salva última linha
        
    columns = ["ITEM", "DESCRIÇÃO", "MATERIAL", "QUANT.", "UND.", "DN (mm)", "dn (mm)", "L (mm)"]
    df = pd.DataFrame(table_data, columns=columns)
    return df

def save_to_excel(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Lista de Materiais", index=False)
        
        # Adicionar formatação ao arquivo Excel
        workbook = writer.book
        worksheet = writer.sheets["Lista de Materiais"]
        
        # Definir estilos de formatação
        header_format = workbook.add_format({'bold': True, 'bg_color': '#9FCF7C', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1, 'align': 'center'})
        
        # Aplicar formatação aos cabeçalhos
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Ajustar largura das colunas e escrever os dados na planilha
        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, 20)  # Ajustando a largura das colunas
        
        for row_num, row_data in enumerate(df.values, start=1):
            for col_num, cell_data in enumerate(row_data):
                worksheet.write(row_num, col_num, cell_data, cell_format)
    
    output.seek(0)
    return output

# Streamlit UI
st.title("Extração de Lista de Materiais do PDF")

pdf_file = st.file_uploader("Carregar um arquivo PDF", type=["pdf"])

if pdf_file:
    st.success("Arquivo carregado com sucesso!")
    df = extract_table_from_pdf(pdf_file)
    
    if not df.empty:
        st.write("### Lista de Materiais Extraída:")
        st.dataframe(df)
        
        # Utilizar o nome do arquivo PDF para nomear o arquivo Excel
        filename = pdf_file.name.replace(".pdf", "")
        
        if st.button("Salvar como Excel"):
            excel_data = save_to_excel(df, filename + ".xlsx")
            st.download_button(label="Baixar Excel", data=excel_data, file_name=filename + ".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Não foi possível extrair a tabela. Verifique se o PDF está no formato correto.")
