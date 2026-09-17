# 💻 Sistema de Inventário de TI em Python

Script automatizado em Python desenvolvido para realizar a coleta detalhada de especificações de hardware e software de máquinas Windows e consolidar os dados em uma planilha Excel (`.xlsx`).

O objetivo deste projeto é otimizar a gestão de ativos de TI através do mapeamento dinâmico e centralizado de dados de infraestrutura.

---

## 📌 Conteúdo do Repositório

### 📁 `.venv`
* Ambiente virtual de execução contendo as dependências isoladas do projeto.

### 📁 `build` & `dist`
* Pastas geradas pelo processo de compilação contendo os arquivos temporários e o executável final (`inventario_ti.exe`).

### 📁 Arquivos Raiz
* **`inventario_ti.py`**: Script principal responsável pela coleta de dados via WMI/psutil e geração do relatório em Excel.
* **`executar_inventario.bat`**: Atalho executável para inicialização rápida do script via terminal.
* **`inventario_ti.spec`**: Arquivo de configuração do PyInstaller para empacotamento do projeto em `.exe`.
* **`inventario.xlsx`**: Planilha consolidada com os dados coletados das máquinas.
* **`.gitignore`**: Regras de exclusão para o versionamento do Git.

---

## 🛠️ Tecnologias Utilizadas

* **Python 3**
* **VS Code** (IDE)
* **Git & GitHub** (Controle de versão)
* **WMI & psutil** (Coleta de métricas de sistema e hardware)
* **openpyxl** (Manipulação de planilhas Excel)

---

## 🚀 Como Executar os Códigos

1. Clone este repositório:
   ```bash
   git clone [https://github.com/lucianospeixoto/ProjetoInventarioPythonTI.git](https://github.com/lucianospeixoto/ProjetoInventarioPythonTI.git)
   cd ProjetoInventarioPythonTI

---

1. Crie e ative o ambiente virtual:
   ```bash
   python -m venv .venv

2. Instale as dependências do projeto:
   ```bash
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt

3. Execute o script principal:
   ```bash
   .\.venv\Scripts\python.exe inventario_ti.py

---

## 💡 Dicas e Soluções de Problemas

1. Geração da Planilha:
    Na primeira execução do script, o arquivo inventario.xlsx será criado automaticamente na raiz do projeto.

2. Erros de Permissão no PowerShell:
    Se ao tentar rodar os comandos surgir algum bloqueio de execução (PSSecurityException), libere a permissão na sessão atual com o comando abaixo:

    ```bash
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass