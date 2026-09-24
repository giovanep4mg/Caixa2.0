# 📊 Gerenciador de Caixa Diário

Um aplicativo desktop leve e intuitivo desenvolvido em **Python** com interface gráfica em **Tkinter** e **EasyGUI**, projetado para automatizar o controle de caixa diário, cálculos de lucros, saldos bancários e o acumulado de moedas e totais.

O programa foi criado para facilitar fechamentos de caixa diários, salvando todos os dados de forma estruturada diretamente em planilhas do **Microsoft Excel (`.xlsx`)**, com formatação profissional de valores monetários.

---

## 🚀 Principais Funcionalidades

- **Criar Nova Planilha (`.xlsx`):** Inicializa uma nova planilha limpa já com todas as colunas de controle padronizadas e formatadas.
- **Lançamentos Contínuos:** Permite selecionar o arquivo uma única vez e realizar múltiplos lançamentos em sequência (perguntando ao final de cada salvamento se deseja lançar um novo dia).
- **Configuração Inicial Automática:** Caso a planilha esteja vazia (ex: 1º dia do mês), o programa solicita os valores iniciais de *Moeda Casa* e *Total Anterior* para garantir a precisão dos cálculos de lucro.
- **Acumulado Inteligente de Moedas:** Soma automaticamente o valor de moedas inserido no dia atual com o montante acumulado do dia anterior.
- **Cálculos Automáticos Integrados:** Calcula automaticamente:
  - Total em Dinheiro (Salão + Notas de R$ 2,00)
  - Total em Bancos (Sicoob, Sumup, Nullbank, MercPago)
  - Soma Total do Caixa
  - Lucro Líquido Diário e Faturamento

---

## 🛠️ Tecnologias e Bibliotecas Utilizadas

O projeto utiliza bibliotecas robustas do ecossistema Python:
* **[Python](https://www.python.org/)** — Linguagem principal.
* **[Pandas](https://pandas.pydata.org/)** — Manipulação e estruturação dos dados em formato tabular.
* **[XlsxWriter](https://xlsxwriter.readthedocs.io/)** — Criação e formatação avançada de arquivos Excel.
* **[Tkinter](https://docs.python.org/3/library/tkinter.html)** — Criação da janela de menu principal.
* **[EasyGUI](https://easygui.sourceforge.net/)'** — Geração de caixas de diálogo e formulários amigáveis para preenchimento rápido.
* **[PyInstaller](https://www.pyinstaller.org/)** — Empacotamento do script em um executável autônomo (`.exe`) para Windows.

---

## 🖥️ Como Usar o Executável

1. Acesse a aba **Actions** (Ações) deste repositório no GitHub.
2. Clique na execução mais recente bem-sucedida (marcada com um ícone verde ✔).
3. Role a página até a seção **Artifacts** (Artefatos) e baixe o arquivo compactado do programa.
4. Descompacte e execute o arquivo `.exe` no seu computador.

### Passo a Passo no Aplicativo:
* **Criar Novo Excel:** Escolha o local e o nome do arquivo para iniciar um novo controle mensal ou diário.
* **Editar / Adicionar Dia no Excel:** Selecione uma planilha já existente. O programa carregará os saldos do último dia cadastrado e abrirá o formulário para preenchimento. Preencha os campos e clique em enviar. Ao final, escolha se deseja lançar outro dia ou retornar ao menu.

---

## 📂 Estrutura das Colunas na Planilha Excel

As planilhas geradas contêm as seguintes colunas calculadas automaticamente:
* `Dia` | `Dinh/Salao` | `Notas2` | `Moeda/Salao` | `Dinh/Casa` | `Moeda/Casa`
* `Gasto/Dia` | `TotalDia` | `Sicoob` | `Sumup` | `Nullbank` | `MercPago`
* `TotalBancos` | `Caixa` | `Casa` | `TotalSoma` | `Lucro` | `TotalAnterior`
