import pandas as pd
from fpdf import FPDF
import math
import os
from datetime import datetime

# Criar estrutura de pastas se não existir
if not os.path.exists("planilhas"):
    os.makedirs("planilhas")
    print("Pasta 'planilhas' criada. Coloque seu arquivo 'faltas.xlsx' dentro dela.")

if not os.path.exists("bilhetes_faltas"):
    os.makedirs("bilhetes_faltas")
    print("Pasta 'bilhetes_faltas' criada.")

# Caminhos dos arquivos
caminho_planilha = os.path.join("planilhas", "faltas.xlsx")
pasta_saida = "bilhetes_faltas"

# Verificar se o arquivo existe
if not os.path.exists(caminho_planilha):
    print(f"ERRO: Arquivo não encontrado em '{caminho_planilha}'")
    print("Por favor, coloque o arquivo 'faltas.xlsx' na pasta 'planilhas'")
    exit()

# Leitura da planilha
try:
    df_geral = pd.read_excel(caminho_planilha, sheet_name="Geral")
    df_bradesco = pd.read_excel(caminho_planilha, sheet_name="Bradesco")
    print("Planilha carregada com sucesso!")
except Exception as e:
    print(f"ERRO ao ler a planilha: {e}")
    exit()

def format_data(row):
    return {
        "paciente": row["Paciente"],
        "convenio": row["Convênio"],
        "data": row["Data"].strftime('%d/%m/%Y'),
        "motivo": row["Motivo da falta"]
    }

# Classe para gerar o PDF com layout específico
class BilhetePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_margins(20, 15, 20)  # Margens: esquerda, topo, direita
    
    def header(self): 
        pass
    
    def footer(self): 
        pass

    def bilhete_bradesco(self, dados):
        # Nome do paciente - Negrito
        self.set_font("Arial", "B", 12)
        self.cell(0, 8, dados['paciente'], ln=True, align='L')
        self.ln(2)
        
        # Convênio
        self.set_font("Arial", "", 11)
        self.cell(0, 6, "BRADESCO SAÚDE", ln=True, align='L')
        self.ln(3)
        
        # Texto principal
        self.set_font("Arial", "", 10)
        texto_principal = f"De acordo com os registros do MedTherapy, o paciente não compareceu à(s) sessão(ões) de {dados['data']} (motivo da falta: {dados['motivo']}). Senhor(a) recepcionista, favor gerar o token e colher a assinatura relativa à falta na próxima sessão."
        
        self.multi_cell(0, 5, texto_principal, align='L')
        self.ln(5)
        
        # Linha separadora pontilhada
        self.draw_dotted_line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(6)
        
        # Token gerado
        self.set_font("Arial", "", 10)
        self.cell(25, 6, "Token gerado?", align='L')
        self.cell(8, 6, "Sim", align='L')
        self.cell(8, 6, "(   )", align='L')
        self.cell(8, 6, "Não", align='L')
        self.cell(8, 6, "(   )", align='L')
        self.ln(8)
        
        self.cell(0, 6, "Motivo (apenas se a resposta anterior for não):", ln=True)
        self.ln(3)
        
        # Primeira linha para motivo
        self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(8)
        
        # Segunda linha para motivo
        self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(8)
        
        # Linha separadora pontilhada
        self.draw_dotted_line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(6)
        
        # Falta assinada
        self.cell(27, 6, "Falta assinada?", align='L')
        self.cell(8, 6, "Sim", align='L')
        self.cell(8, 6, "(   )", align='L')
        self.cell(8, 6, "Não", align='L')
        self.cell(8, 6, "(   )", align='L')
        self.ln(8)
        
        self.cell(0, 6, "Motivo (apenas se a resposta anterior for não):", ln=True)
        self.ln(3)
        
        # Primeira linha para motivo da falta assinada
        self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(8)
        
        # Segunda linha para motivo da falta assinada
        self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(10)

    def bilhete_padrao(self, dados):
        # Nome do paciente - Negrito
        self.set_font("Arial", "B", 11)
        self.cell(0, 6, dados['paciente'], ln=True, align='L')
        self.ln(1)
        
        # Convênio
        self.set_font("Arial", "", 11)
        self.cell(0, 6, dados['convenio'], ln=True, align='L')
        self.ln(2)
        
        # Texto principal
        self.set_font("Arial", "", 10)
        texto_principal = f"De acordo com os registros do MedTherapy, o paciente não compareceu à(s) sessão(ões) de {dados['data']}. (motivo da falta: {dados['motivo']}). Senhor(a) recepcionista, favor colher a assinatura relativa à falta na próxima sessão."
        
        self.multi_cell(0, 4.5, texto_principal, align='L')
        self.ln(3)
        
        # Linha separadora pontilhada
        self.draw_dotted_line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(4)
        
        # Falta assinada
        self.cell(27, 5, "Falta assinada?", align='L')
        self.cell(8, 5, "Sim", align='L')
        self.cell(8, 5, "(   )", align='L')
        self.cell(8, 5, "Não", align='L')
        self.cell(8, 5, "(   )", align='L')
        self.ln(8)
        
        self.cell(0, 5, "Motivo (apenas se a resposta anterior for não):", ln=True)
        self.ln(6)
        
        # Linha para motivo
        self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
        self.ln(8)
        
        # Linhas para assinatura
        for i in range(3):
            self.line(self.get_x(), self.get_y(), self.get_x() + 170, self.get_y())
            self.ln(8)

    def draw_dotted_line(self, x1, y1, x2, y2):
        """Desenha uma linha pontilhada"""
        self.set_line_width(0.3)
        current_x = x1
        while current_x < x2:
            line_end = min(current_x + 2, x2)
            self.line(current_x, y1, line_end, y1)
            current_x += 4

# Gerar nome do arquivo com data e hora
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
nome_arquivo = f"bilhetes_faltas_{timestamp}.pdf"
caminho_saida = os.path.join(pasta_saida, nome_arquivo)

# Gerar PDF
pdf = BilhetePDF()
pdf.set_auto_page_break(auto=False)

# ---- Bradesco: 2 por página (formatação especial)
bradesco_linhas = [format_data(row) for _, row in df_bradesco.iterrows()]
print(f"Gerando {len(bradesco_linhas)} bilhetes do Bradesco...")

for i in range(0, len(bradesco_linhas), 2):
    pdf.add_page()
    
    for j in range(2):
        if i + j < len(bradesco_linhas):
            y_position = 20 + j * 140  # Posição Y para cada bilhete (espaçamento de 140 para mais espaço)
            pdf.set_y(y_position)
            pdf.bilhete_bradesco(bradesco_linhas[i + j])

# ---- Outros: 3 por página
outros_df = df_geral[df_geral["Convênio"].str.upper().str.contains("BRADESCO") == False]
outros_linhas = [format_data(row) for _, row in outros_df.iterrows()]
print(f"Gerando {len(outros_linhas)} bilhetes de outros convênios...")

for i in range(0, len(outros_linhas), 3):
    pdf.add_page()
    
    for j in range(3):
        if i + j < len(outros_linhas):
            y_position = 15 + j * 95  # Posição Y para cada bilhete (espaçamento de 95)
            pdf.set_y(y_position)
            pdf.bilhete_padrao(outros_linhas[i + j])

# Salvar
try:
    pdf.output(caminho_saida)
    print(f"PDF gerado com sucesso!")
    print(f"Arquivo salvo em: {caminho_saida}")
    print(f"Total de bilhetes gerados: {len(bradesco_linhas) + len(outros_linhas)}")
except Exception as e:
    print(f"ERRO ao salvar o PDF: {e}")