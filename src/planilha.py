
import pandas as pd
import numpy as np
import re
import tkinter as tk
import sys
import subprocess
import os
import time
from tkinter import filedialog, messagebox
from collections import defaultdict
from decimal import Decimal, InvalidOperation


# Importa a biblioteca Pillow para lidar com imagens (necessária para .jpg)
# Certifique-se de que a biblioteca Pillow esteja instalada:
# pip install Pillow
try:
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showerror("Erro de Dependência", 
                         "A biblioteca 'Pillow' (PIL) não foi encontrada.\n"
                         "Por favor, instale-a abrindo o terminal e executando:\n"
                         "pip install Pillow")
    Image, ImageTk = None, None
    
# Certifique-se de que a biblioteca BeautifulSoup esteja instalada:
# pip install beautifulsoup4
try:
    from bs4 import BeautifulSoup
except ImportError:
    messagebox.showerror("Erro de Dependência", 
                         "A biblioteca 'BeautifulSoup4' não foi encontrada.\n"
                         "Por favor, instale-a abrindo o terminal e executando:\n"
                         "pip install beautifulsoup4")
    BeautifulSoup = None


# --- REGEX ---
# Padrão para extrair 'Aquisicao' ou 'Pagamento' e o número da nota fiscal
padrao_movimentacao = re.compile(r'(AQUISICAO|PAGAMENTO).*?(\d+)', re.IGNORECASE)

# --- FUNÇÕES AUXILIARES ---
def parse_valor_br(s: str) -> Decimal:
    """Converte uma string de valor em formato brasileiro para Decimal."""
    try:
        if isinstance(s, (int, float)):
            return Decimal(str(s))
        s = str(s).replace(".", "").replace(",", ".")
        return Decimal(s)
    except InvalidOperation:
        return Decimal("0.00")

def fmt_br(d: Decimal) -> str:
    """Formata um Decimal para string de valor em formato brasileiro."""
    if not isinstance(d, Decimal):
        d = Decimal(str(d))
    return f"{d:.2f}".replace(".", ",")

def processar_planilha_xlsx(caminho_entrada, pasta_saida_relatorios):
    """
    Processa um único arquivo .xlsx, extrai dados de movimentação e gera relatórios.
    caminho_entrada: o caminho do arquivo .xlsx de entrada.
    pasta_saida_relatorios: o caminho da pasta onde o relatório .txt será salvo.
    """
    try:
        # Lê o arquivo completo sem cabeçalho para ter controle total
        df_bruto = pd.read_excel(caminho_entrada, header=None, engine='openpyxl')
        
        # Encontra a linha de cabeçalho
        row_with_headers = -1
        for i, row in df_bruto.iterrows():
            row_str = [str(x).upper() for x in row]
            if 'DÉBITO' in row_str and 'CRÉDITO' in row_str:
                row_with_headers = i
                break
                
        if row_with_headers == -1:
            print(f"Aviso: Não foi possível encontrar a linha de cabeçalho em '{os.path.basename(caminho_entrada)}'.")
            return

        # Encontra os índices de todas as colunas de interesse
        header_row_data = df_bruto.iloc[row_with_headers]
        
        col_index_data = header_row_data[header_row_data.astype(str).str.contains('DATA', na=False, case=False)].first_valid_index()
        col_index_historico = header_row_data[header_row_data.astype(str).str.contains('CONTRAPARTIDA/HISTÓRICO', na=False, case=False)].first_valid_index()
        col_index_debito = header_row_data[header_row_data.astype(str).str.contains('DÉBITO', na=False, case=False)].first_valid_index()
        col_index_credito = header_row_data[header_row_data.astype(str).str.contains('CRÉDITO', na=False, case=False)].first_valid_index()
        col_index_saldo = header_row_data[header_row_data.astype(str).str.contains('SALDO-EXERCÍCIO', na=False, case=False)].first_valid_index()

        if any(idx is None for idx in [col_index_data, col_index_historico, col_index_debito, col_index_credito]):
            print(f"Aviso: Uma ou mais colunas essenciais não foram encontradas em '{os.path.basename(caminho_entrada)}'.")
            return
        


        # Extrai o saldo anterior procurando pela descrição na coluna de histórico
        saldoAnterior_val = Decimal("0.00")
        if col_index_saldo is not None:
            # Procura a linha com "SALDO ANTERIOR"
            linha_saldo_anterior = df_bruto.iloc[row_with_headers:].astype(str).apply(
                lambda row: any("SALDO ANTERIOR" in str(cell).upper() for cell in row), axis=1
            )
            
            if linha_saldo_anterior.any():
                indice_saldo = linha_saldo_anterior[linha_saldo_anterior].index[0]
                try:
                    saldo_anterior_bruto = df_bruto.iloc[indice_saldo, col_index_saldo]
                    saldoAnterior_val = parse_valor_br(saldo_anterior_bruto)
                    saldoAnterior_val = (saldoAnterior_val * -1 if saldoAnterior_val < 0 else saldoAnterior_val)

                    print(f"Saldo Anterior extraído: {fmt_br(saldoAnterior_val)}")
                except (IndexError, KeyError, InvalidOperation):
                    print("Aviso: Não foi possível extrair o Saldo Anterior.")
            else:
                print("Aviso: 'SALDO ANTERIOR' não encontrado na planilha.")
        


        # Seleciona os dados a partir da linha seguinte à do cabeçalho
        df_final = df_bruto.iloc[row_with_headers + 1:, [col_index_data, col_index_historico, col_index_debito, col_index_credito, col_index_saldo]].copy()
        df_final.columns = ['Data', 'Texto_Completo', 'Débito', 'Crédito', 'Saldo']

        # Converte 'Data' para o formato correto e remove linhas inválidas
        df_final['Data'] = pd.to_datetime(df_final['Data'], errors='coerce')
        # 
        df_final.dropna(subset=['Data'], inplace=True)
        # 

        # Extrai Descrição e Número da coluna de texto
        extraido = df_final['Texto_Completo'].astype(str).str.extract(padrao_movimentacao)
        df_final['Descrição'] = extraido[0]
        df_final['Numero'] = extraido[1]

        # Converte as colunas de valores para numérico
        df_final['Débito'] = pd.to_numeric(df_final['Débito'], errors='coerce').fillna(0)
        df_final['Crédito'] = pd.to_numeric(df_final['Crédito'], errors='coerce').fillna(0)
        df_final['Saldo'] = pd.to_numeric(df_final['Saldo'], errors='coerce').fillna(0)
        
        # Remove linhas que não tenham a descrição ou o número
        # 
        df_final.dropna(subset=['Descrição', 'Numero'], inplace=True)
        # 

        # Reseta o índice para começar do zero
        df_final = df_final.reset_index(drop=True)

        # --- ENGENHARIA E CÁLCULOS (lógica similar à do PDF) ---
        notas = defaultdict(lambda: {"credito": Decimal("0.00"), "debito": Decimal("0.00")})
        relatorio = []
        
        somaSomenteDebito = 0

        for index, row in df_final.iterrows():
            nf = str(row['Numero'])
            debito_val = parse_valor_br(row['Débito'])
            credito_val = parse_valor_br(row['Crédito'])
            saldo_val = parse_valor_br(row['Saldo'])

            # print(f"NF  {nf}  -  {saldoAnterior_val}")

            # A lógica é simplificada aqui para somar diretamente os valores
            notas[nf]["debito"] += debito_val
            notas[nf]["credito"] += credito_val

        # GERAR RELATÓRIO .txt
        for nf, valores in notas.items():
            credito, debito = valores["credito"], valores["debito"]
            if credito == 0 and debito == 0:
                continue

            diferenca = credito - debito
            status = ""
            if credito > 0 and debito == 0:
                status = "Sem pagamento registrado"
            elif debito > 0 and credito == 0:
                status = "Sem aquisição registrada"

                somaSomenteDebito += debito 

            elif abs(diferenca) < Decimal("0.01"):
                status = "OK"
            else:
                status = f"Diferença {fmt_br(diferenca)}"

            relatorio.append(f"NF {nf} -> Crédito: {fmt_br(credito)} | Débito: {fmt_br(debito)} | {status}")
        
        print(f"Soma Débito {somaSomenteDebito}")
        print(f"Saldo Anterior {saldoAnterior_val}")
        
        if somaSomenteDebito > 0:
            print("Cálculo Saldo Anterior")

            diferenca = somaSomenteDebito - saldoAnterior_val 

            if abs(diferenca) < Decimal("0.01"):
                status = f"| Saldo Anterior OK | Saldo Anterior {fmt_br(saldoAnterior_val)}    Débito Sem Aquisição Registrada {fmt_br(somaSomenteDebito)}"
            else:
                status = f"| Saldo Anterior Diferença {fmt_br(diferenca)} | Saldo Anterior {fmt_br(saldoAnterior_val)}    Débito Sem Aquisição Registrada {fmt_br(somaSomenteDebito)}"
        else:
            status = f"| Saldo Anterior OK | Saldo Anterior {fmt_br(saldoAnterior_val)}    Não existe Aquisição Registrada"

        relatorio.append(f"{status}")

        if relatorio:
            nome_base = os.path.splitext(os.path.basename(caminho_entrada))[0]
            # O relatório .txt será salvo na pasta de saída escolhida
            caminho_saida_txt = os.path.join(pasta_saida_relatorios, f"{nome_base}_relatorio.txt")
            with open(caminho_saida_txt, "w", encoding="utf-8") as f:
                f.write("\n".join(relatorio))
            print(f"Relatório de '{os.path.basename(caminho_entrada)}' salvo em: {caminho_saida_txt}")

        # GERAR PLANILHA FINAL
        nome_base = os.path.splitext(os.path.basename(caminho_entrada))[0]
        # A planilha final será salva na mesma pasta do arquivo de entrada
        caminho_saida_lancamentos_xlsx = os.path.join(os.path.dirname(caminho_entrada), f"{nome_base}_lancamentos.xlsx")
        df_final.to_excel(caminho_saida_lancamentos_xlsx, index=False)
        print(f"Dados processados de '{os.path.basename(caminho_entrada)}' salvos em: {caminho_saida_lancamentos_xlsx}")

    except Exception as e:
        print(f"Ocorreu um erro ao processar '{os.path.basename(caminho_entrada)}': {e}")
        

# --- INTERFACE (Tkinter) ---
def escolher_pasta(entry_widget):
    """Abre uma caixa de diálogo para escolher uma pasta e preenche o widget de entrada."""
    pasta = filedialog.askdirectory()
    if pasta:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, pasta)

def find_libreoffice_path():
    """Tenta encontrar o caminho do executável do LibreOffice em locais comuns."""
    common_paths = [
        os.path.join(os.getenv("PROGRAMFILES", "C:\\Program Files"), "LibreOffice\\program\\soffice.exe"),
        os.path.join(os.getenv("PROGRAMFILES(X86)", "C:\\Program Files (x86)"), "LibreOffice\\program\\soffice.exe"),
        "/usr/bin/libreoffice", # Para sistemas Linux
        "/Applications/LibreOffice.app/Contents/MacOS/soffice" # Para sistemas macOS
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path
    return None

def executar():
    """Função principal que orquestra a conversão e o processamento de planilhas."""
    pasta_entrada = pasta_entry.get()
    pasta_saida = saida_entry.get()
    
    if not os.path.isdir(pasta_entrada):
        messagebox.showerror("Erro", "Selecione uma pasta de ENTRADA válida.")
        return
    if not os.path.isdir(pasta_saida):
        messagebox.showerror("Erro", "Selecione uma pasta de SAÍDA válida.")
        return

    arquivos_encontrados = [f for f in os.listdir(pasta_entrada) if f.lower().endswith(('.xls', '.xlsx'))]
    if not arquivos_encontrados:
        messagebox.showinfo("Aviso", "Nenhum arquivo .xls ou .xlsx encontrado na pasta de entrada.")
        return

    print("Iniciando o processamento...")
    for arquivo in arquivos_encontrados:
        caminho_completo_entrada = os.path.join(pasta_entrada, arquivo)
        nome_base, extensao = os.path.splitext(arquivo)

        caminho_para_processar = caminho_completo_entrada
        
        # Se for um arquivo .xls, tenta convertê-lo primeiro
        if extensao.lower() == '.xls':
            print(f"\nDetectado arquivo .xls: '{arquivo}'. Iniciando a conversão...")
            
            # O arquivo convertido será salvo na mesma pasta de entrada
            caminho_convertido = os.path.join(pasta_entrada, f"{nome_base}.xlsx")
            
            # Tenta encontrar o LibreOffice antes de tentar a conversão
            soffice_path = find_libreoffice_path()
            if not soffice_path:
                messagebox.showerror("Erro de Conversão", "LibreOffice não encontrado. Certifique-se de que está instalado.")
                return

            try:
                # O --outdir para a conversão deve ser a pasta de entrada
                comando_libreoffice = f'"{soffice_path}" --headless --convert-to xlsx --outdir "{pasta_entrada}" "{caminho_completo_entrada}"'
                
                subprocess.run(comando_libreoffice, shell=True, check=True)
                
                # Espere um pouco para o LibreOffice terminar a conversão
                time.sleep(2)
                
                if os.path.exists(caminho_convertido):
                    caminho_para_processar = caminho_convertido
                    print(f"Conversão concluída. Arquivo salvo como '{caminho_convertido}'.")
                else:
                    print(f"Erro: Conversão de '{arquivo}' falhou ou o arquivo de saída não foi encontrado.")
                    continue
            except subprocess.CalledProcessError as e:
                print(f"Erro de subprocesso ao tentar converter '{arquivo}': {e}")
                continue
        
        # Processa o arquivo (original .xlsx ou o recém-convertido)
        # Passa o caminho do arquivo de entrada e a pasta de saída de relatórios para a função
        processar_planilha_xlsx(caminho_para_processar, pasta_saida)
    
    messagebox.showinfo("Processamento concluído", f"Verifique a pasta de saída para os relatórios e a pasta de entrada para as planilhas processadas.")

def make_image_transparent(image):
    """
    Converte pixels brancos (ou muito claros) para transparentes.
    """
    if not image:
        return None
    image = image.convert("RGBA")
    # Converte a imagem para o modo RGB, pois PyInstaller pode ter problemas com CMYK
    if image.mode == 'CMYK':
        image = image.convert('RGB')
    image = image.convert("RGBA")
    
    datas = image.getdata()
    
    newData = []
    for item in datas:
        # Troca pixels brancos (ou quase brancos) por transparentes
        if item[0] > 240 and item[1] > 240 and item[2] > 240:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)
    
    image.putdata(newData)
    return image

# --- CRIAÇÃO DA JANELA TKINTER ---
root = tk.Tk()
root.title("Análise de Balancete licenciado para G.A.B.CONTABILIDADE")
root.resizable(False, False)

# Altera o ícone da janela para a imagem fornecida (necessita de 'Pillow')
# Certifique-se de que o arquivo 'icon.jpg' está na mesma pasta que o script.

# Define o caminho base para encontrar arquivos, compatível com PyInstaller
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

try:
    if Image and ImageTk:
        icon_path = os.path.join(base_path, "icon.jpg")
        if os.path.exists(icon_path):
            icon_image = Image.open(icon_path)

            # Torna o fundo branco da imagem transparente
            icon_image_transparent = make_image_transparent(icon_image)

            # Redimensiona a imagem para o novo tamanho de ícone (60x60)
            icon_image_resized = icon_image.resize((60, 60), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(icon_image_resized)
            root.iconphoto(False, photo)
        else:
            print(f"Aviso: Arquivo de ícone '{icon_path}' não encontrado.")
except Exception as e:
    print(f"Erro ao tentar definir o ícone: {e}")

# Pasta de entrada
tk.Label(root, text="Pasta de Planilhas (.xls/.xlsx):").grid(row=0, column=0, padx=10, pady=10, sticky="e")
pasta_entry = tk.Entry(root, width=50)
pasta_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="📁", command=lambda: escolher_pasta(pasta_entry)).grid(row=0, column=2, padx=10, pady=10)

# Pasta de saída
tk.Label(root, text="Pasta para salvar relatórios:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
saida_entry = tk.Entry(root, width=50)
saida_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="📁", command=lambda: escolher_pasta(saida_entry)).grid(row=1, column=2, padx=10, pady=10)

# Adiciona um novo rótulo para o texto adicional
# tk.Label(root, text="\U0001F4DA G.A.B.CONTABILIDADE").grid(row=2, column=0, pady=(10, 5))

# Ícone de livro
icone_livro = "  \U0001F4DA"

# Label para o ícone (fonte grande)
tk.Label(root, text=icone_livro, font=("Arial", 20)).grid(row=2, column=0, pady=(10, 5), sticky="w") # 'sticky="e"' alinha à direita

# Label para o texto (fonte menor)
tk.Label(root, text="G.A.B. CONTABILIDADE", font=("Arial", 8)).grid(row=2, column=0, pady=(12, 5), sticky="e") # 'sticky="w"' alinha à esquerda


# Juntos eles ficam um ao lado do outro na mesma linha 2


# Adiciona o texto antes do botão "Processar"
try:
    image_path = os.path.join(base_path, 'icon.jpg')
    if Image and ImageTk and os.path.exists(image_path):
        # Abre a imagem usando PIL
        pil_image = Image.open(image_path)
        # Torna o fundo branco da imagem transparente e redimensiona
        pil_image_transparent = make_image_transparent(pil_image)
        pil_image_transparent = pil_image_transparent.resize((60, 60), Image.Resampling.LANCZOS)
        
        # Converte a imagem PIL para um objeto PhotoImage que o Tkinter pode usar
        tk_image = ImageTk.PhotoImage(pil_image_transparent)

        # Cria um Frame para agrupar a imagem e o texto
        frame_dev = tk.Frame(root)
        frame_dev.grid(row=2, column=0, pady=(10, 5))

        frame_dev = tk.Frame(root)        
        frame_dev.grid(row=2, column=1, pady=(10, 5))

        # O fundo do frame para combinar com o da janela
        frame_dev.config(bg=root['bg']) 

        # Cria o rótulo para a imagem e a exibe no frame
        image_label = tk.Label(frame_dev, image=tk_image)
        image_label.pack(side=tk.LEFT, padx=(0, 5))
        image_label.config(bg=root['bg']) # O fundo do label para combinar com o da janela


    # Cria o rótulo com o texto, agora no mesmo frame
    text_label = tk.Label(frame_dev, text="Desenvolvido por Denis Menegon - \u260e (19) 99493-4477", font=("Helvetica", 10))
    text_label.pack(side=tk.LEFT)
    
except FileNotFoundError:
    # Caso a imagem não seja encontrada, exibe um rótulo de erro
    tk.Label(root, text="Erro: A imagem 'icon.jpg' não foi encontrada.", fg="red").grid(row=2, column=1, pady=(10, 5))
except Exception as e:
    tk.Label(root, text=f"Erro ao carregar a imagem: {e}", fg="red").grid(row=2, column=1, pady=(10, 5))


# Botão processar
tk.Button(root, text="Processar", command=executar, bg="#3956b6", fg="white").grid(row=3, column=1, pady=(5, 20), sticky="e")
root.mainloop()


# -*- coding: utf-8 -*-

# import pandas as pd
# import numpy as np
# import re
# import tkinter as tk
# from tkinter import filedialog, messagebox
# import subprocess
# import os
# import time
# from collections import defaultdict
# from decimal import Decimal, InvalidOperation

# # Importa a biblioteca Pillow para lidar com imagens (necessária para .jpg)
# # Certifique-se de que a biblioteca Pillow esteja instalada:
# # pip install Pillow
# try:
#     from PIL import Image, ImageTk
# except ImportError:
#     messagebox.showerror("Erro de Dependência", 
#                          "A biblioteca 'Pillow' (PIL) não foi encontrada.\n"
#                          "Por favor, instale-a abrindo o terminal e executando:\n"
#                          "pip install Pillow")
#     Image, ImageTk = None, None
    
# # Certifique-se de que a biblioteca BeautifulSoup esteja instalada:
# # pip install beautifulsoup4
# try:
#     from bs4 import BeautifulSoup
# except ImportError:
#     messagebox.showerror("Erro de Dependência", 
#                          "A biblioteca 'BeautifulSoup4' não foi encontrada.\n"
#                          "Por favor, instale-a abrindo o terminal e executando:\n"
#                          "pip install beautifulsoup4")
#     BeautifulSoup = None

# # --- REGEX ---
# # Padrão para extrair 'Aquisicao' ou 'Pagamento' e o número da nota fiscal
# padrao_movimentacao = re.compile(r'(AQUISICAO|PAGAMENTO).*?(\d+)', re.IGNORECASE)
# # Padrão para extrair o número da nota fiscal de um histórico
# padrao_nota_fiscal = re.compile(r'NOTA FISCAL (\d+)', re.IGNORECASE)
# # Padrão para limpar valores numéricos
# padrao_limpeza_valor = re.compile(r'[^\d,.]')

# # --- FUNÇÕES AUXILIARES ---
# def parse_valor_br(s: str) -> Decimal:
#     """Converte uma string de valor em formato brasileiro para Decimal."""
#     try:
#         if isinstance(s, (int, float)):
#             return Decimal(str(s))
#         # Verifica se o valor é vazio ou NaN antes de tentar substituir
#         if pd.isna(s) or s.strip() == '':
#             return Decimal("0.00")
#         s = str(s).replace(".", "").replace(",", ".")
#         return Decimal(s)
#     except (InvalidOperation, TypeError):
#         return Decimal("0.00")

# def fmt_br(d: Decimal) -> str:
#     """Formata um Decimal para string de valor em formato brasileiro."""
#     if not isinstance(d, Decimal):
#         d = Decimal(str(d))
#     return f"{d:.2f}".replace(".", ",")

# def processar_dados_e_gerar_relatorio(df_final, caminho_entrada, pasta_saida_relatorios):
#     """
#     Função centralizada para processar o DataFrame e gerar os relatórios.
#     Isso evita duplicação de código para .xlsx e .htm.
#     """
    
#     # Reseta o índice para começar do zero
#     df_final = df_final.reset_index(drop=True)

#     # --- ENGENHARIA E CÁLCULOS (lógica similar à do PDF) ---
#     notas = defaultdict(lambda: {"credito": Decimal("0.00"), "debito": Decimal("0.00")})
#     relatorio = []
    
#     # Apenas processa as linhas com número de Nota Fiscal para o relatório .txt
#     df_relatorio_txt = df_final.dropna(subset=['Numero']).copy()
        
#     for index, row in df_relatorio_txt.iterrows():
#         nf = str(row['Numero'])
#         # Apenas processa se um número de nota fiscal foi encontrado
#         if nf and nf != "nan":
#             # Converte as colunas de valores para numérico antes de usar
#             debito_val = parse_valor_br(row['Débito'])
#             credito_val = parse_valor_br(row['Crédito'])

#             # A lógica é simplificada aqui para somar diretamente os valores
#             notas[nf]["debito"] += debito_val
#             notas[nf]["credito"] += credito_val

#     # GERAR RELATÓRIO .txt
#     if notas:
#         for nf, valores in notas.items():
#             credito, debito = valores["credito"], valores["debito"]
#             if credito == 0 and debito == 0:
#                 continue

#             diferenca = credito - debito
#             status = ""
#             if credito > 0 and debito == 0:
#                 status = "Sem pagamento registrado"
#             elif debito > 0 and credito == 0:
#                 status = "Sem aquisição registrada"
#             elif abs(diferenca) < Decimal("0.01"):
#                 status = "OK"
#             else:
#                 status = f"Diferença {fmt_br(diferenca)}"

#             relatorio.append(f"NF {nf} -> Crédito: {fmt_br(credito)} | Débito: {fmt_br(debito)} | {status}")
    
#         if relatorio:
#             nome_base = os.path.splitext(os.path.basename(caminho_entrada))[0]
#             # O relatório .txt será salvo na pasta de saída escolhida
#             caminho_saida_txt = os.path.join(pasta_saida_relatorios, f"{nome_base}_relatorio.txt")
#             with open(caminho_saida_txt, "w", encoding="utf-8") as f:
#                 f.write("\n".join(relatorio))
#             print(f"Relatório de '{os.path.basename(caminho_entrada)}' salvo em: {caminho_saida_txt}")

#     # GERAR PLANILHA FINAL (AGORA INCLUINDO TODAS AS LINHAS)
#     nome_base = os.path.splitext(os.path.basename(caminho_entrada))[0]
#     # A planilha final será salva na mesma pasta do arquivo de entrada
#     caminho_saida_lancamentos_xlsx = os.path.join(os.path.dirname(caminho_entrada), f"{nome_base}_lancamentos.xlsx")
    
#     # Define a ordem final das colunas antes de salvar
#     df_final = df_final[['Data', 'Texto_Completo', 'Débito', 'Crédito', 'Descrição', 'Numero']]
    
#     # Salva a planilha com todas as linhas que foram lidas e processadas
#     df_final.to_excel(caminho_saida_lancamentos_xlsx, index=False)
#     print(f"Dados processados de '{os.path.basename(caminho_entrada)}' salvos em: {caminho_saida_lancamentos_xlsx}")


# def processar_planilha_xlsx(caminho_entrada, pasta_saida_relatorios):
#     """
#     Processa um único arquivo .xlsx, extrai dados de movimentação e gera relatórios.
#     caminho_entrada: o caminho do arquivo .xlsx de entrada.
#     pasta_saida_relatorios: o caminho da pasta onde o relatório .txt será salvo.
#     """
#     try:
#         # Lê o arquivo completo sem cabeçalho para ter controle total
#         df_bruto = pd.read_excel(caminho_entrada, header=None, engine='openpyxl')
        
#         # Encontra a linha de cabeçalho
#         row_with_headers = -1
#         for i, row in df_bruto.iterrows():
#             row_str = [str(x).upper() for x in row]
#             if 'DÉBITO' in row_str and 'CRÉDITO' in row_str:
#                 row_with_headers = i
#                 break
                
#         if row_with_headers == -1:
#             print(f"Aviso: Não foi possível encontrar a linha de cabeçalho em '{os.path.basename(caminho_entrada)}'.")
#             return

#         # Encontra os índices de todas as colunas de interesse
#         header_row_data = df_bruto.iloc[row_with_headers]
        
#         col_index_data = header_row_data[header_row_data.astype(str).str.contains('DATA', na=False, case=False)].first_valid_index()
#         col_index_historico = header_row_data[header_row_data.astype(str).str.contains('CONTRAPARTIDA/HISTÓRICO', na=False, case=False)].first_valid_index()
#         col_index_debito = header_row_data[header_row_data.astype(str).str.contains('DÉBITO', na=False, case=False)].first_valid_index()
#         col_index_credito = header_row_data[header_row_data.astype(str).str.contains('CRÉDITO', na=False, case=False)].first_valid_index()
        
#         if any(idx is None for idx in [col_index_data, col_index_historico, col_index_debito, col_index_credito]):
#             print(f"Aviso: Uma ou mais colunas essenciais não foram encontradas em '{os.path.basename(caminho_entrada)}'.")
#             return

#         # Seleciona os dados a partir da linha seguinte à do cabeçalho
#         df_final = df_bruto.iloc[row_with_headers + 1:, [col_index_data, col_index_historico, col_index_debito, col_index_credito]].copy()
#         df_final.columns = ['Data', 'Texto_Completo', 'Débito', 'Crédito']

#         # Converte 'Data' para o formato correto e remove linhas inválidas
#         df_final['Data'] = pd.to_datetime(df_final['Data'], errors='coerce')
#         df_final.dropna(subset=['Data'], inplace=True)
        
#         # Extrai Descrição e Número da coluna de texto
#         extraido = df_final['Texto_Completo'].astype(str).str.extract(padrao_movimentacao)
#         df_final['Descrição'] = extraido[0]
#         df_final['Numero'] = extraido[1]
        
#         # Converte as colunas de valores para numérico
#         df_final['Débito'] = pd.to_numeric(df_final['Débito'], errors='coerce').fillna(0)
#         df_final['Crédito'] = pd.to_numeric(df_final['Crédito'], errors='coerce').fillna(0)
        
#         # Chama a função centralizada para continuar o processamento
#         processar_dados_e_gerar_relatorio(df_final, caminho_entrada, pasta_saida_relatorios)

#     except Exception as e:
#         print(f"Ocorreu um erro ao processar '{os.path.basename(caminho_entrada)}': {e}")


# def processar_planilha_htm(caminho_entrada, pasta_saida_relatorios):
#     """
#     Processa um arquivo .htm, extrai dados de movimentação e gera relatórios.
#     caminho_entrada: o caminho do arquivo .htm de entrada.
#     pasta_saida_relatorios: o caminho da pasta onde o relatório .txt será salvo.
#     """
#     if BeautifulSoup is None:
#         return # Impede a execução se a dependência não foi encontrada
    
#     # Tenta abrir o arquivo com diferentes codificações
#     codificacoes_tentativa = ['utf-8', 'latin-1', 'cp1252']
#     html_content = None
#     for encoding in codificacoes_tentativa:
#         try:
#             with open(caminho_entrada, 'r', encoding=encoding) as f:
#                 html_content = f.read()
#             # Se a leitura for bem-sucedida, saia do loop
#             break
#         except UnicodeDecodeError:
#             continue
    
#     if html_content is None:
#         print(f"Erro: Não foi possível decodificar o arquivo '{os.path.basename(caminho_entrada)}' com as codificações padrão.")
#         return

#     try:
#         soup = BeautifulSoup(html_content, 'html.parser')
        
#         # Encontra a tabela principal
#         table = soup.find('table')
#         if not table:
#             print(f"Aviso: Nenhuma tabela encontrada no arquivo HTML '{os.path.basename(caminho_entrada)}'.")
#             return

#         data = []
#         rows = table.find_all('tr')
        
#         # Itera sobre as linhas de dados, ignorando a linha de cabeçalho
#         print("Iniciando a extração dos dados da tabela HTML...")
#         for i, row in enumerate(rows):
#             tds = row.find_all('td')
#             # Verifica se a linha tem o número esperado de colunas
#             # Assumimos que a tabela possui pelo menos 9 colunas
#             if len(tds) >= 9:
#                 try:
#                     # Usamos find_all(text=True) para obter todo o texto, mesmo que esteja em tags aninhadas
#                     data_val = ''.join(tds[0].find_all(text=True)).strip()
#                     historico_val = ''.join(tds[3].find_all(text=True)).strip().replace('\n', ' ')
#                     debito_val = ''.join(tds[7].find_all(text=True)).strip()
#                     credito_val = ''.join(tds[8].find_all(text=True)).strip()
                    
#                     # Log para debug
#                     print(f"Linha {i+1}: Débito lido -> '{debito_val}', Crédito lido -> '{credito_val}'")

#                     # Valida se a linha tem dados que parecem ser uma transação
#                     if data_val and (debito_val or credito_val):
#                         data.append({
#                             'Data': data_val,
#                             'Texto_Completo': historico_val,
#                             'Débito': debito_val,
#                             'Crédito': credito_val
#                         })
#                 except IndexError:
#                     # Ignora linhas que não têm o número de colunas esperado
#                     continue
        
#         if not data:
#             print(f"Aviso: Nenhuma linha de dados válida encontrada em '{os.path.basename(caminho_entrada)}'.")
#             return

#         # Cria o DataFrame a partir dos dados extraídos do HTML
#         df_final = pd.DataFrame(data)
#         print(f"DataFrame inicial criado com {len(df_final)} linhas.")

#         # Converte 'Data' para o formato correto e remove linhas inválidas
#         df_final['Data'] = pd.to_datetime(df_final['Data'], errors='coerce', dayfirst=True)
#         df_final.dropna(subset=['Data'], inplace=True)
        
#         # Extrai Descrição e Numero
#         extraido = df_final['Texto_Completo'].astype(str).str.extract(padrao_movimentacao)
#         df_final['Descrição'] = extraido[0]
#         df_final['Numero'] = extraido[1]

#         # Limpa os valores usando o novo padrão regex mais robusto
#         df_final['Débito'] = df_final['Débito'].astype(str).str.replace(padrao_limpeza_valor, '', regex=True)
#         df_final['Crédito'] = df_final['Crédito'].astype(str).str.replace(padrao_limpeza_valor, '', regex=True)

#         # Converte as colunas de valores para numérico
#         df_final['Débito'] = pd.to_numeric(df_final['Débito'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)
#         df_final['Crédito'] = pd.to_numeric(df_final['Crédito'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)
        
#         print(f"DataFrame final antes de gerar o relatório tem {len(df_final)} linhas e as seguintes colunas:\n{df_final.columns}")

#         # Chama a função centralizada para continuar o processamento
#         processar_dados_e_gerar_relatorio(df_final, caminho_entrada, pasta_saida_relatorios)
        
#     except Exception as e:
#         print(f"Ocorreu um erro ao processar '{os.path.basename(caminho_entrada)}': {e}")
    

# # --- INTERFACE (Tkinter) ---
# def escolher_pasta(entry_widget):
#     """Abre uma caixa de diálogo para escolher uma pasta e preenche o widget de entrada."""
#     pasta = filedialog.askdirectory()
#     if pasta:
#         entry_widget.delete(0, tk.END)
#         entry_widget.insert(0, pasta)

# def find_libreoffice_path():
#     """Tenta encontrar o caminho do executável do LibreOffice em locais comuns."""
#     common_paths = [
#         os.path.join(os.getenv("PROGRAMFILES", "C:\\Program Files"), "LibreOffice\\program\\soffice.exe"),
#         os.path.join(os.getenv("PROGRAMFILES(X86)", "C:\\Program Files (x86)"), "LibreOffice\\program\\soffice.exe"),
#         "/usr/bin/libreoffice", # Para sistemas Linux
#         "/Applications/LibreOffice.app/Contents/MacOS/soffice" # Para sistemas macOS
#     ]
#     for path in common_paths:
#         if os.path.exists(path):
#             return path
#     return None

# def executar():
#     """Função principal que orquestra a conversão e o processamento de planilhas."""
#     pasta_entrada = pasta_entry.get()
#     pasta_saida = saida_entry.get()
    
#     if not os.path.isdir(pasta_entrada):
#         messagebox.showerror("Erro", "Selecione uma pasta de ENTRADA válida.")
#         return
#     if not os.path.isdir(pasta_saida):
#         messagebox.showerror("Erro", "Selecione uma pasta de SAÍDA válida.")
#         return

#     arquivos_encontrados = [f for f in os.listdir(pasta_entrada) if f.lower().endswith(('.xls', '.xlsx', '.htm', '.html'))]
#     if not arquivos_encontrados:
#         messagebox.showinfo("Aviso", "Nenhum arquivo .xls, .xlsx, .htm ou .html encontrado na pasta de entrada.")
#         return

#     print("Iniciando o processamento...")
#     for arquivo in arquivos_encontrados:
#         caminho_completo_entrada = os.path.join(pasta_entrada, arquivo)
#         nome_base, extensao = os.path.splitext(arquivo)

#         if extensao.lower() in ['.xls', '.xlsx']:
#             # Lógica para arquivos de planilha (existente)
#             caminho_para_processar = caminho_completo_entrada
            
#             if extensao.lower() == '.xls':
#                 print(f"\nDetectado arquivo .xls: '{arquivo}'. Iniciando a conversão...")
                
#                 caminho_convertido = os.path.join(pasta_entrada, f"{nome_base}.xlsx")
                
#                 soffice_path = find_libreoffice_path()
#                 if not soffice_path:
#                     messagebox.showerror("Erro de Conversão", "LibreOffice não encontrado. Certifique-se de que está instalado.")
#                     return

#                 try:
#                     comando_libreoffice = f'"{soffice_path}" --headless --convert-to xlsx --outdir "{pasta_entrada}" "{caminho_completo_entrada}"'
#                     subprocess.run(comando_libreoffice, shell=True, check=True)
#                     time.sleep(2)
                    
#                     if os.path.exists(caminho_convertido):
#                         caminho_para_processar = caminho_convertido
#                         print(f"Conversão concluída. Arquivo salvo como '{caminho_convertido}'.")
#                     else:
#                         print(f"Erro: Conversão de '{arquivo}' falhou ou o arquivo de saída não foi encontrado.")
#                         continue
#                 except subprocess.CalledProcessError as e:
#                     print(f"Erro de subprocesso ao tentar converter '{arquivo}': {e}")
#                     continue
            
#             processar_planilha_xlsx(caminho_para_processar, pasta_saida)
            
#         elif extensao.lower() in ['.htm', '.html']:
#             # Nova lógica para arquivos HTML
#             print(f"\nDetectado arquivo HTML: '{arquivo}'. Iniciando o processamento...")
#             processar_planilha_htm(caminho_completo_entrada, pasta_saida)
    
#     messagebox.showinfo("Processamento concluído", "Verifique a pasta de saída para os relatórios e a pasta de entrada para as planilhas processadas.")


# def make_image_transparent(image):
#     """
#     Converte pixels brancos (ou muito claros) para transparentes.
#     """
#     if not image:
#         return None
#     image = image.convert("RGBA")
#     # Converte a imagem para o modo RGB, pois PyInstaller pode ter problemas com CMYK
#     if image.mode == 'CMYK':
#         image = image.convert('RGB')
#     image = image.convert("RGBA")
    
#     datas = image.getdata()
    
#     newData = []
#     for item in datas:
#         # Troca pixels brancos (ou quase brancos) por transparentes
#         if item[0] > 240 and item[1] > 240 and item[2] > 240:
#             newData.append((255, 255, 255, 0))
#         else:
#             newData.append(item)
    
#     image.putdata(newData)
#     return image

# # --- CRIAÇÃO DA JANELA TKINTER ---
# root = tk.Tk()
# root.title("Análise de Balancete licenciado para G.A.B.CONTABILIDADE")
# root.resizable(False, False)

# # Altera o ícone da janela para a imagem fornecida (necessita de 'Pillow')
# # Certifique-se de que o arquivo 'icon.jpg' está na mesma pasta que o script.
# try:
#     if Image and ImageTk:
#         icon_path = "icon.jpg" # Use o nome do seu arquivo de imagem
#         if os.path.exists(icon_path):
#             icon_image = Image.open(icon_path)

#             # Torna o fundo branco da imagem transparente
#             icon_image_transparent = make_image_transparent(icon_image)

#             # Redimensiona a imagem para o novo tamanho de ícone (60x60)
#             icon_image_resized = icon_image.resize((60, 60), Image.Resampling.LANCZOS)
#             photo = ImageTk.PhotoImage(icon_image_resized)
#             root.iconphoto(False, photo)
#         else:
#             print(f"Aviso: Arquivo de ícone '{icon_path}' não encontrado.")
# except Exception as e:
#     print(f"Erro ao tentar definir o ícone: {e}")

# # Pasta de entrada
# tk.Label(root, text="Pasta de Planilhas (.xls/.xlsx/.htm):").grid(row=0, column=0, padx=10, pady=10, sticky="e")
# pasta_entry = tk.Entry(root, width=50)
# pasta_entry.grid(row=0, column=1, padx=10, pady=10)
# tk.Button(root, text="📁", command=lambda: escolher_pasta(pasta_entry)).grid(row=0, column=2, padx=10, pady=10)

# # Pasta de saída
# tk.Label(root, text="Pasta para salvar relatórios:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
# saida_entry = tk.Entry(root, width=50)
# saida_entry.grid(row=1, column=1, padx=10, pady=10)
# tk.Button(root, text="📁", command=lambda: escolher_pasta(saida_entry)).grid(row=1, column=2, padx=10, pady=10)

# # Adiciona um novo rótulo para o texto adicional
# tk.Label(root, text="G.A.B.CONTABILIDADE").grid(row=2, column=0, pady=(10, 5))


# # Adiciona o texto antes do botão "Processar"
# try:
#     image_path = 'icon.jpg'
#     if Image and ImageTk and os.path.exists(image_path):
#         # Abre a imagem usando PIL
#         pil_image = Image.open(image_path)
#         # Torna o fundo branco da imagem transparente e redimensiona
#         pil_image_transparent = make_image_transparent(pil_image)
#         pil_image_transparent = pil_image_transparent.resize((60, 60), Image.Resampling.LANCZOS)
        
#         # Converte a imagem PIL para um objeto PhotoImage que o Tkinter pode usar
#         tk_image = ImageTk.PhotoImage(pil_image_transparent)

#         # Cria um Frame para agrupar a imagem e o texto
#         frame_dev = tk.Frame(root)
#         frame_dev.grid(row=2, column=0, pady=(10, 5))

#         frame_dev = tk.Frame(root)        
#         frame_dev.grid(row=2, column=1, pady=(10, 5))

#         # O fundo do frame para combinar com o da janela
#         frame_dev.config(bg=root['bg']) 

#         # Cria o rótulo para a imagem e a exibe no frame
#         image_label = tk.Label(frame_dev, image=tk_image)
#         image_label.pack(side=tk.LEFT, padx=(0, 5))
#         image_label.config(bg=root['bg']) # O fundo do label para combinar com o da janela


#     # Cria o rótulo com o texto, agora no mesmo frame
#     text_label = tk.Label(frame_dev, text="        Desenvolvido por Denis Menegon - \u260e (19) 99493-4477", font=("Helvetica", 10))
#     text_label.pack(side=tk.LEFT)
    
# except FileNotFoundError:
#     # Caso a imagem não seja encontrada, exibe um rótulo de erro
#     tk.Label(root, text="Erro: A imagem 'icon.jpg' não foi encontrada.", fg="red").grid(row=2, column=1, pady=(10, 5))
# except Exception as e:
#     tk.Label(root, text=f"Erro ao carregar a imagem: {e}", fg="red").grid(row=2, column=1, pady=(10, 5))


# # Botão processar
# tk.Button(root, text="Processar", command=executar, bg="#3956b6", fg="white").grid(row=3, column=1, pady=(5, 20))

# root.mainloop()
