import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import easygui as eg
import xlsxwriter

def converter_valor(valor_str):
    """Trata string vazia, substitui vírgula por ponto e converte para float com segurança."""
    if valor_str is None or str(valor_str).strip() == "":
        return 0.0
    tratado = str(valor_str).strip().replace(',', '.')
    return float(tratado)

def criar_arquivo_excel():
    """Cria um novo arquivo Excel do zero escolhendo o local e nome pelo Windows."""
    nome_arquivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx", 
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Salvar Novo Arquivo Excel"
    )
    if nome_arquivo:
        try:
            dados = {
                'Dia': [], 'Dinh/Salao': [], 'Notas2': [], 'Moeda/Salao': [], 'Dinh/Casa': [],
                'Moeda/Casa': [], 'Gasto/Dia': [], 'TotalDia': [], 'Sicoob': [], 'Sumup': [],
                'Nullbank': [], 'MercPago': [], 'TotalBancos': [], 'Caixa': [], 'Casa': [],
                'TotalSoma': [], 'Lucro': [], 'TotalAnterior': []
            }
            df = pd.DataFrame(dados)
            
            writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
            df.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            formato_numero = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('B:S', 15, formato_numero)
            writer.close()
            
            messagebox.showinfo("Sucesso", f"Arquivo '{os.path.basename(nome_arquivo)}' criado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar arquivo: {e}")

def editar_arquivo_excel():
    """Abre um arquivo existente, lida com dados iniciais se vazio, preenche o formulário e atualiza."""
    nome_arquivo = filedialog.askopenfilename(
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Selecione o Arquivo Excel do Caixa"
    )
    if not nome_arquivo:
        return

    try:
        df = pd.read_excel(nome_arquivo)
    except FileNotFoundError:
        messagebox.showerror("Erro", f"O arquivo não foi encontrado.")
        return
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir arquivo: {e}")
        return

    tem_dados = not df.empty and len(df) > 0
    
    # Se a planilha estiver vazia, pede os valores anteriores manualmente (ex: 1º dia do mês)
    if not tem_dados:
        msg_inicial = "A planilha está vazia (1º Lançamento).\nInsira os saldos anteriores paraIsso é um ajuste super importante para o fluxo do seu aplicativo de caixa! Para lidar com o primeiro dia do mês — ou qualquer momento em que você precise inserir um saldo inicial ou dados acumulados anteriores —, precisamos modificar a lógica do seu código Python (provavelmente no script que gera o executável com o PyInstaller).

Aqui está como você pode estruturar essa adição no seu código:

### 1. Atualizar a Entrada de Dados (Interface ou Terminal)
Se você estiver usando `tkinter` ou `easygui`, pode criar campos específicos (ou uma janela de diálogo condicional) que só aparecem se for o fechamento/abertura do mês, ou deixar os campos sempre disponíveis de forma opcional.

Exemplo conceitual de como capturar esses valores em Python:

```python
import easygui as eg


def obter_dados_iniciais():
    # Pergunta se deseja inserir dados anteriores (ideal para o 1º dia do mês)
    continuar = eg.ccbox(
        "É o primeiro dia do mês ou deseja inserir dados anteriores?",
        "Dados Anteriores",
        choices=("Sim", "Não"),
    )

    valor_moedas_antigas = 0.0
    valor_total_anterior = 0.0

    if continuar:
        msg = "Insira os dados do período anterior:"
        title = "Configuração Inicial do Mês"
        fieldNames = [
            "Valor acumulado em moedas anterior (R$):",
            "Valor total anterior (R$):",
        ]
        fieldValues = eg.multbox(msg, title, fieldNames)

        if fieldValues:
            try:
                valor_moedas_antigas = float(fieldValues[0].replace(",", "."))
                valor_total_anterior = float(fieldValues[1].replace(",", "."))
            except ValueError:
                eg.msgbox(
                    "Erro: Digite apenas números válidos!", "Erro de Formato"
                )
                return obter_dados_iniciais()  # Tenta novamente se houver erro

    return valor_moedas_antigas, valor_total_anterior
