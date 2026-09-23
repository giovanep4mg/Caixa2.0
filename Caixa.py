import os
import pandas as pd
import easygui as eg
import tkinter as tk
from tkinter import messagebox

NOME_ARQUIVO = "caixa_teste.xlsx"

def carregar_dados():
    if os.path.exists(NOME_ARQUIVO):
        try:
            return pd.read_excel(NOME_ARQUIVO)
        except Exception:
            pass
    
    dados = {
        'Dia': [], 'Dinh/Salao': [], 'Notas2': [], 'Moeda/Salao': [], 'Dinh/Casa': [],
        'Moeda/Casa': [], 'Gasto/Dia': [], 'TotalDia': [], 'Sicoob': [], 'Sumup': [],
        'Nullbank': [], 'MercPago': [], 'TotalBancos': [], 'Caixa': [], 'Casa': [],
        'TotalSoma': [], 'Lucro': [], 'TotalAnterior': []
    }
    return pd.DataFrame(dados)

def converter_valor(valor_str):
    if valor_str is None or str(valor_str).strip() == "":
        return 0.0
    tratado = str(valor_str).strip().replace(',', '.')
    return float(tratado)

def interface_grafica():
    df = carregar_dados()
    
    ultimo_sicoob = df['Sicoob'].iloc[-1] if not df.empty and 'Sicoob' in df.columns else 0.0
    ultimo_sumup = df['Sumup'].iloc[-1] if not df.empty and 'Sumup' in df.columns else 0.0
    ultimo_nullbank = df['Nullbank'].iloc[-1] if not df.empty and 'Nullbank' in df.columns else 0.0
    ultimo_mercPago = df['MercPago'].iloc[-1] if not df.empty and 'MercPago' in df.columns else 0.0
    ultimo_moedaCasa = df['Moeda/Casa'].iloc[-1] if not df.empty and 'Moeda/Casa' in df.columns else 0.0
    ultimo_totalSoma = df['TotalSoma'].iloc[-1] if not df.empty and 'TotalSoma' in df.columns else 0.0

    msg = "Preencha os dados do Caixa Diário (Use ponto ou vírgula para decimais):"
    title = "Controle de Caixa"
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
            return

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
        writer = pd.ExcelWriter(NOME_ARQUIVO, engine='xlsxwriter')
        df.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        formato_numero = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column('B:S', 15, formato_numero)
        writer.close()
        
        eg.msgbox(f"✅ Dados do dia {dia} salvos com sucesso no arquivo '{NOME_ARQUIVO}'!", "Sucesso")
    except Exception as e:
        eg.msgbox(f"❌ Erro ao salvar o arquivo Excel: {e}", "Erro de Salvamento")

if __name__ == "__main__":
    interface_grafica()
