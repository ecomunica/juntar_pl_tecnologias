# juntar_pl_tecnologias
### How-To: Instalação e Utilização do Script de Processamento de Planilhas com Tkinter e ttkbootstrap

---

### 1. **Pré-requisitos**

Antes de iniciar, verifique se você tem o **Python 3** instalado no seu sistema. Para verificar, abra um terminal ou prompt de comando e digite:

```bash
python --version
```

Se o Python não estiver instalado, você pode baixá-lo e instalá-lo [aqui](https://www.python.org/downloads/).

### 2. **Instalando as bibliotecas necessárias**

O script requer algumas bibliotecas Python adicionais. As principais são:

- **ttkbootstrap** para estilizar a interface gráfica.
- **pandas** para manipulação de dados.
- **pyexcel** e **pyexcel-ods3** para leitura de arquivos `.ods` (OpenDocument Spreadsheet).

#### Passo 1: Atualizar o `pip`

Atualize o `pip` (gerenciador de pacotes do Python) para garantir que as instalações sejam bem-sucedidas. Execute:

```bash
python -m pip install --upgrade pip
```

#### Passo 2: Instalar as dependências

Use o comando abaixo para instalar todas as bibliotecas necessárias:

```bash
pip install pandas pyexcel pyexcel-ods3 ttkbootstrap
```

### 3. **Estrutura do Script**

O script processa arquivos `.ods` em um diretório especificado pelo usuário e gera um arquivo Excel consolidado com as informações de equipes, colaboradores, funções e tecnologias. 
Além disso, a interface gráfica inclui uma barra de progresso que indica visualmente que o processamento está em andamento.

---

### 4. **Como executar o script**

#### Passo 1: Salvar o script Python

1. Abra um editor de texto como **Notepad**, **VS Code**, **Sublime Text**, ou qualquer editor de sua escolha.
2. Copie o código Python abaixo e salve-o como **`processa_planilhas.py`**.

```python
import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
from tkinter import filedialog
import os
import pandas as pd
import pyexcel as pe
import threading

# Função para processar as planilhas do diretório selecionado e gerar 'dados_coletados.xlsx'
def processar_planilhas(diretorio, progress_bar, texto_resultado):
    arquivos = [f for f in os.listdir(diretorio) if f.endswith('.ods')]
    dados_coletados = []

    def buscar_valor_ao_lado(aba, palavra_chave):
        """Busca o valor ao lado de uma célula que contém uma palavra-chave."""
        for r in range(aba.number_of_rows()):
            for c in range(aba.number_of_columns()):
                valor_celula = str(aba.cell_value(r, c)).strip().lower()  # Padroniza a busca
                if palavra_chave in valor_celula:
                    # Tenta pegar o valor à direita
                    try:
                        return str(aba.cell_value(r, c + 1)).strip()  # Valor ao lado (coluna à direita)
                    except IndexError:
                        return None
        return None

    # Função para processar cada arquivo .ods
    def processar_arquivo(arquivo):
        caminho_arquivo = os.path.join(diretorio, arquivo)
        livro = pe.get_book(file_name=caminho_arquivo)
        aba_nomes_normalizados = {nome.strip().lower(): nome for nome in livro.sheet_names()}
        
        for nome_normalizado, nome_original in aba_nomes_normalizados.items():
            try:
                aba = livro.sheet_by_name(nome_original)
                equipe = buscar_valor_ao_lado(aba, "equipe") or "N/A"
                colaborador = buscar_valor_ao_lado(aba, "colaborador") or "N/A"
                funcao_principal = buscar_valor_ao_lado(aba, "função principal") or "N/A"
                
                tecnologias = []
                for i in range(10, 20):  # B11 -> index 10 até B20 -> index 19
                    tecnologia = str(aba.cell_value(i, 1)).strip()
                    if tecnologia and tecnologia.lower() != 'selecione':
                        tecnologias.append(tecnologia)
                
                # Se não houver tecnologias válidas, adicionar "Vazio"
                if not tecnologias:
                    dados_coletados.append([equipe, colaborador, funcao_principal, "Vazio"])
                else:
                    # Adicionar as tecnologias encontradas
                    for tecnologia in tecnologias:
                        dados_coletados.append([equipe, colaborador, funcao_principal, tecnologia])
            except Exception as e:
                with open('log_erros.txt', 'a') as log_file:
                    log_file.write(f"Erro ao acessar aba '{nome_original}' no arquivo '{arquivo}': {e}\n")

    # Limpar o campo de resultados antes de iniciar
    texto_resultado.delete(1.0, tk.END)
    
    # Iniciar a barra de progresso indeterminada
    progress_bar.start()

    for arquivo in arquivos:
        processar_arquivo(arquivo)

    # Parar a barra de progresso indeterminada
    progress_bar.stop()

    # Após processamento, salvar o arquivo 'dados_coletados.xlsx'
    if dados_coletados:
        df = pd.DataFrame(dados_coletados, columns=['Equipe', 'Colaborador', 'Função Principal', 'Tecnologia'])
        df.to_excel('dados_coletados.xlsx', index=False)
        texto_resultado.insert(tk.END, "Arquivo 'dados_coletados.xlsx' gerado com sucesso!\n")
        
        # Contagem de colaboradores por equipe e extração dos primeiros nomes
        contagem = df.groupby('Equipe').agg({'Colaborador': lambda x: list(set(x))})
        
        for equipe, colaboradores in contagem.iterrows():
            # Extrair o primeiro nome de cada colaborador
            primeiros_nomes = [colab.split()[0] for colab in colaboradores['Colaborador']]
            quantidade = len(primeiros_nomes)
            # Exibir a quantidade e os primeiros nomes
            texto_resultado.insert(tk.END, f"{equipe} - {quantidade} colaboradores ({', '.join(primeiros_nomes)})\n")

# Função para selecionar o diretório
def selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    if diretorio:
        label_diretorio.config(text=f"Diretório selecionado: {diretorio}")
        botao_iniciar.config(state=tk.NORMAL)  # Habilitar o botão de iniciar
        return diretorio
    return None

# Função para iniciar o processamento em uma nova thread
def iniciar_processamento():
    diretorio = label_diretorio.cget("text").replace("Diretório selecionado: ", "")
    thread = threading.Thread(target=processar_planilhas, args=(diretorio, progress_bar, texto_resultado))
    thread.start()

# Criar a janela principal com ttkbootstrap
janela = tb.Window(themename="darkly")
janela.title("Contagem de Colaboradores por Equipe")
janela.geometry("600x400")

# Torna os widgets responsivos ao redimensionar a janela
janela.rowconfigure(0, weight=1)
janela.rowconfigure(1, weight=1)
janela.rowconfigure(2, weight=1)
janela.rowconfigure(3, weight=3)  # O campo de texto cresce mais
janela.columnconfigure(0, weight=1)

# Label para mostrar o diretório selecionado
label_diretorio = ttk.Label(janela, text="Nenhum diretório selecionado")
label_diretorio.grid(row=0, column=0, pady=10, sticky="ew")

# Botão para selecionar o diretório com largura fixa e padding
botao_diretorio = ttk.Button(janela, text="Selecionar Diretório", command=selecionar_diretorio)
botao_diretorio.grid(row=1, column=0, pady=10, padx=80, ipadx=10, sticky="ew")

# Botão de iniciar processamento (inicialmente desabilitado)
botao_iniciar = ttk.Button(janela, text="Iniciar Processamento", command=iniciar_processamento, state=tk.DISABLED)
botao_iniciar.grid(row=2, column=0, pady=10, padx=80, ipadx=10, sticky="ew")

# Barra de progresso personalizada com tema "success"
progress_bar = ttk.Progressbar(janela, orient="horizontal", mode="indeterminate", bootstyle="success-striped")
progress_bar.grid(row=3, column=0, pady=10, padx=80, sticky="ew")

# Campo de texto para mostrar os resultados
texto_resultado = tk.Text(janela, height=10, width=60)
texto_resultado.grid(row=4, column=0, pady=10, padx=10, sticky="nsew")

# Rodar a aplicação
janela.mainloop()
```

#### Passo 2: Executar o script

1. Abra um terminal ou prompt de comando.
2. Navegue até o diretório onde o script foi salvo.
3. Execute o script com o comando:

```bash
python processa_planilhas.py
```

### 5. **Uso do Script**

#### Passo 1: Abrir o programa
Após rodar o comando `python processa_planilhas.py`, a interface gráfica será aberta.

#### Passo 2: Selecionar o diretório com as planilhas
1. Clique no botão **"Selecionar Diretório"** para escolher a pasta onde estão os arquivos `.ods`.
   - Uma janela de seleção será aberta para navegar e escolher o diretório com as planilhas.
   - O caminho do diretório escolhido será exibido na interface.

#### Passo 3: Iniciar o processamento
1. Após selecionar o diretório, o botão **"Iniciar Processamento"** será habilitado.
2. Clique em **"Iniciar Processamento"** para começar o processamento das planilhas.
   - A barra de progresso começará a se mover para indicar que o processamento está em andamento.
   - Durante o processamento, a interface permanecerá responsiva e exibirá uma barra de progresso indeterminada.
   
#### Passo 4: Visualizar os resultados
1. Ao final do processamento, o arquivo **`dados_coletados.xlsx`** será gerado no mesmo diretório em que o script foi executado.
2. Na própria interface, será exibido:
   - Uma mensagem de confirmação: `"Arquivo 'dados_coletados.xlsx' gerado com sucesso!"`.
   - A quantidade de colaboradores por equipe, incluindo os primeiros nomes dos colaboradores, no formato:  
     **Equipe1 - 17 colaboradores (Adriana, Allan, Carla, ...)**
   
3. Se qualquer erro ocorrer durante o processamento, ele será registrado em um arquivo de log chamado **`log_erros.txt`**, que será salvo no mesmo diretório de execução.

---

### 6. **Estrutura do Arquivo Gerado**

O arquivo **`dados_coletados.xlsx`** conterá os dados consolidados das planilhas `.ods`. As colunas principais do arquivo são:

- **Equipe**: O nome da equipe do colaborador.
- **Colaborador**: O nome completo do colaborador.
- **Função Principal**: A função principal desempenhada pelo colaborador.
- **Tecnologia**: As tecnologias nas quais o colaborador tem experiência, ou "Vazio" caso não haja informações.

---

### 7. **Solução de Problemas**

Caso encontre problemas ao executar o script, considere as seguintes soluções:

1. **Dependências não instaladas**:
   - Se você receber um erro indicando que uma biblioteca está ausente, certifique-se de que as dependências foram instaladas corretamente usando o comando:
   ```bash
   pip install pandas pyexcel pyexcel-ods3 ttkbootstrap
   ```

2. **Erro ao processar planilhas**:
   - Caso ocorra algum erro ao acessar as planilhas, ele será registrado no arquivo **`log_erros.txt`**. Verifique este arquivo para mais informações.

3. **Interface gráfica não abre**:
   - Se a interface gráfica não abrir, verifique se o Tkinter está instalado corretamente. No terminal, você pode executar:
   ```bash
   python -m tkinter
   ```
   - Se o Tkinter não estiver instalado, você pode instalá-lo dependendo do seu sistema operacional:
     - **Linux (Ubuntu/Debian)**:
       ```bash
       sudo apt-get install python3-tk
       ```
     - **Windows** e **macOS** geralmente já incluem o Tkinter com a instalação do Python.

---

### 8. **Considerações Finais**

---




