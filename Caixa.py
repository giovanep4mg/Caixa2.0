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
    """Abre um arquivo existente, preenche os dados de forma unificada e atualiza a planilha."""
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

    # Pega valores anteriores para preencher os campos automaticamente
    tem_dados = not df.empty and len(df) > 0
    
    ultimo_sicoob = df['Sicoob'].iloc[-1] if tem_dados and 'Sicoob' in df.columns and pd.notna(df['Sicoob'].iloc[-1]) else 0.0
    ultimo_sumup = df['Sumup'].iloc[-1] if tem_dados and 'Sumup' in df.columns and pd.notna(df['Sumup'].iloc[-1]) else 0.0
    ultimo_nullbank = df['Nullbank'].iloc[-1] if tem_dados and 'Nullbank' in df.columns and pd.notna(df['Nullbank'].iloc[-1]) else 0.0
    ultimo_mercPago = df['MercPago'].iloc[-1] if tem_dados and 'MercPago' in df.columns and pd.notna(df['MercPago'].iloc[-1]) else 0.0
    ultimo_moedaCasa = df['Moeda/Casa'].iloc[-1] if tem_dados and 'Moeda/Casa' in df.columns and pd.notna(df['Moeda/Casa'].iloc[-1]) else 0.0
    ultimo_totalSoma = df['TotalSoma'].iloc[-1] if tem_dados and 'TotalSoma' in df.columns and pd.notna(df['TotalSoma'].iloc[-1]) else 0.0

    msg = "Preencha os dados do Caixa Diário (Use ponto ou vírgula para decimais):"
    title = "Controle de Caixa - Lançamento"
    fieldNames = [
        "Dia", "Dinheiro Salão", "Notas de 2", "Moeda Salão", 
        "Dinheiro Casa", "Moeda Casa", "Gasto do Dia", 
        "Sicoob", "Sumup", "Nullbank", "MercPago"
    ]
    
    fieldValues = [
        "", "", "", "", "", 
        str(ultimo_moedaCasa), "", 
        str(ultimo_sicoob), str(ultimo_sumup), 
        str(ultimo_nullbank), str(ultimo_mercPago)
    ]

    while True:
        resposta = eg.multenterbox(msg, title, fieldNames, fieldValues)
        if resposta is None:
            return  # Cancelado pelo usuário

        try:
            dia = int(converter_valor(resposta[0]))
            dinhSalao = converter_valor(resposta[1])
            notas2 = converter_valor(resposta[2])
            moedaSalao = converter_valor(resposta[3])
            dinhCasa = converter_valor(resposta[4])
            moedaCasa_informado = converter_valor(resposta[5])
            gastoDia = converter_valor(resposta[6])
            
            sicoob = converter_valor(resposta[7])
            sumup = converter_valor(resposta[8])
            nullbank = converter_valor(resposta[9])
            mercPago = converter_valor(resposta[10])
            break
        except ValueError:
            eg.msgbox("❌ Erro em algum valor numérico! Verifique se digitou letras ou caracteres inválidos (use ponto ou vírgula para decimais).", "Erro de Formato")
            fieldValues = resposta

    # Lógica de cálculos
    casa = dinhCasa
    caixa = dinhSalao + notas2
    totalbancos = sicoob + sumup + nullbank + mercPago

    moedaCasa_atual = moedaCasa_informado
    totalsoma = casa + caixa + totalbancos + moedaSalao + moedaCasa_atual

    total_anterior = ultimo_totalSoma
    novo_lucro = totalsoma - total_anterior
    totalDia = novo_lucro + gastoDia

    nova_linha = {
        'Dia': dia, 'Dinh/Salao': dinhSalao, 'Notas2': notas2, 'Moeda/Salao': moedaSalao,
        'Dinh/Casa': dinhCasa, 'Moeda/Casa': moedaCasa_atual, 'Gasto/Dia': gastoDia,
        'Sicoob': sicoob, 'Sumup': sumup, 'Nullbank': nullbank, 'MercPago': mercPago,
        'TotalBancos': totalbancos, 'Casa': casa, 'Caixa': caixa, 'TotalSoma': totalsoma,
        'Lucro': novo_lucro, 'TotalDia': totalDia, 'TotalAnterior': total_anterior
    }

    df_novo = pd.DataFrame([nova_linha])
    df = pd.concat([df, df_novo], ignore_index=True)

    try:
        writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
        df.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        formato_numero = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column('B:S', 15, formato_numero)
        writer.close()
        
        messagebox.showinfo("Sucesso", f"Dados do dia {dia} salvos e arquivo atualizado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel (Verifique se ele está aberto no Excel): {e}")

def criar_janela():
    janela = tk.Tk()
    janela.title("Gerenciador de Caixa")
    janela.geometry("500x350")
    janela.config(bg="#f0f0f0")

    # Título visual na janela
    titulo = tk.Label(janela, text="Controle de Caixa Diário", font=("Arial", 16, "bold"), bg="#f0f0f0")
    titulo.pack(pady=20)

    botao_novo = tk.Button(janela, text="Criar Novo Excel", command=criar_arquivo_excel, width=25, height=2, font=("Arial", 11), bg="#4CAF50", fg="white")
    botao_novo.pack(pady=10)

    botao_editar = tk.Button(janela, text="Editar / Adicionar Dia no Excel", command=editar_arquivo_excel, width=25, height=2, font=("Arial", 11), bg="#2196F3", fg="white")
    botao_editar.pack(pady=10)

    janela.mainloop()

if __name__ == "__main__":
    criar_janela()
