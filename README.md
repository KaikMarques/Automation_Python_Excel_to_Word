Aqui está a documentação técnica completa e profissional para o seu projeto, já atualizada com a nova funcionalidade de "Salvar Como" (escolher nome e local).

Você pode criar um arquivo chamado `README.md` na raiz do seu projeto e colar o conteúdo abaixo.

---

# 📑 Automação de Contratos e Relatórios (Excel → Word)

> **Ferramenta de RPA (Robotic Process Automation) para geração massiva e consolidada de documentos.**

## 🎯 Visão Geral

Este software foi desenvolvido para otimizar o fluxo de trabalho de departamentos que lidam com a criação repetitiva de contratos ou relatórios baseados em dados tabelados. A aplicação oferece uma **interface gráfica (GUI) moderna**, desenvolvida em Python com Tkinter, que permite ao usuário transformar linhas de uma planilha Excel em documentos Word formatados e consolidados.

A versão atual (**v2.0**) implementa a liberdade total de salvamento, permitindo ao usuário definir o nome do arquivo final e o diretório de destino em uma única etapa.

---

## ✨ Funcionalidades

* **Interface Intuitiva:** Design limpo importado do Figma, com feedback visual de seleção.
* **Seleção de Fonte de Dados:** Importação de planilhas `.xlsx`.
* **Templating Dinâmico:** Preenchimento de modelos `.docx` utilizando tags Jinja2.
* **Flexibilidade de Saída:** Funcionalidade "Salvar Como..." para definir nome personalizado e local do relatório.
* **Merge Automático:** Consolidação de múltiplos documentos gerados em um único arquivo mestre.
* **Limpeza Inteligente:** Remoção automática de arquivos temporários após o processamento.
* **Tratamento de Erros:** Sistema de logs visuais (pop-ups) para alertar sobre falhas de leitura ou execução.

---

## 🛠️ Tecnologias e Dependências

O projeto foi construído utilizando **Python 3.12+**. As seguintes bibliotecas são necessárias:

| Biblioteca | Função |
| --- | --- |
| `tkinter` | Interface Gráfica (Nativa do Python). |
| `pandas` | Manipulação e leitura da base de dados Excel. |
| `docxtpl` | Motor de template para Word (substituição de variáveis). |
| `docxcompose` | Unificação (merge) de documentos Word. |
| `openpyxl` | Engine para leitura de arquivos `.xlsx`. |

Para instalar as dependências, execute:

```bash
pip install pandas docxtpl docxcompose openpyxl python-docx pyinstaller

```

---

## 📂 Estrutura de Arquivos Obrigatória

Para o correto funcionamento do código fonte (modo desenvolvimento) e compilação, a estrutura de pastas deve ser respeitada:

```text
Projeto/
├── gui.py                  # Código fonte principal
├── README.md               # Este arquivo
└── assets/                 # Recursos gráficos
    └── frame0/             # Imagens exportadas do Figma
        ├── image_1.png
        ├── button_1.png
        └── ...

```

---

## 📋 Especificação dos Dados de Entrada

Para que a automação funcione, os arquivos de entrada devem seguir estritamente o padrão abaixo:

### 1. Base de Dados (Excel)

O arquivo `.xlsx` deve conter uma aba chamada **`Matriz_Aceitacao`** com os seguintes cabeçalhos exatos:

| Nome da Empresa | Atividade da Empresa | Funcionários | Gasto Anual | Faturamento Anual |
| --- | --- | --- | --- | --- |
| *Texto* | *Texto* | *Número* | *Número* | *Número* |

### 2. Modelo de Documento (Word)

O arquivo `.docx` (Template) deve conter as variáveis (tags) onde os dados serão inseridos. A formatação (negrito, fonte, cor) aplicada à tag será mantida no texto final.

* `{{nome_empresa}}`
* `{{atividade}}`
* `{{funcionarios}}`
* `{{gasto_anual}}`
* `{{faturamento}}`

---

## 🚀 Guia de Utilização

1. **Execução:** Inicie a aplicação (`gui.py` ou `gui.exe`).
2. **Passo 1 (Excel):** Clique no botão correspondente para selecionar a planilha de dados.
* *Feedback:* O nome do arquivo aparecerá em **Verde**.


3. **Passo 2 (Modelo):** Selecione o arquivo `.docx` que servirá de template.
* *Feedback:* O nome do arquivo aparecerá em **Azul**.


4. **Passo 3 (Salvar Como):** Clique no botão para definir onde o arquivo será salvo e qual será seu nome (ex: `Relatorio_Final_Outubro.docx`).
* *Feedback:* O nome escolhido aparecerá em **Laranja**.


5. **Processamento:** Clique no botão **"GERAR RELATÓRIO"**.
* O sistema irá processar linha por linha, criar os documentos na memória, unificá-los e salvar o arquivo final.
* Uma mensagem de "Sucesso" será exibida ao final.



> **Dica de Usabilidade:** Caso selecione um arquivo incorreto, clique com o **Botão Direito do Mouse** sobre o botão para limpar a seleção.

---

## 📦 Como Criar o Executável (.exe)

Para distribuir a ferramenta para a equipe sem a necessidade de instalar Python, utilize o **PyInstaller**. O código já está preparado com a função `resource_path` para suportar arquivos estáticos.

1. Abra o terminal na pasta onde está o arquivo `gui.py` e a pasta `assets`.
2. Execute o comando:

```bash
pyinstaller --noconsole --onefile --windowed --add-data "assets;assets" gui.py

```

* **--noconsole:** Não abre a tela preta de terminal.
* **--onefile:** Gera um único arquivo executável.
* **--add-data:** Embuti a pasta de imagens dentro do executável.

O arquivo final estará na pasta `dist/`.

---

## 📞 Suporte e Manutenção

**Desenvolvedor:** [Seu Nome]
**Status:** Produção (Stable)
**Contato:** [Seu Email ou Teams]
