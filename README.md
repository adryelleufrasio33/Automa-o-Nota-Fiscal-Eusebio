# Automação de Extração de Dados de Nota Fiscal – Prefeitura de Eusébio

Este projeto é uma **automação em Python** para leitura de **notas fiscais eletrônicas** da Prefeitura de Eusébio (CE).  
A aplicação percorre uma pasta no Windows, coleta os arquivos de nota fiscal e gera automaticamente uma **planilha no formato LibreOffice Calc (.ods)** consolidando informações relevantes como:

- Descrição  
- Valor  
- ISS  
- Outros campos configurados no código  

Essa automação trouxe **grande ganho de produtividade**, reduzindo tarefas manuais e **economizando tempo do setor administrativo**, que antes precisava preencher planilhas manualmente.

---

## 🚀 Funcionalidades

- Leitura automática de arquivos de notas fiscais dentro de uma pasta especificada.  
- Extração de campos pré-definidos (descrição, valor, ISS etc.).  
- Geração de planilha **ODS** (LibreOffice Calc) com os dados organizados.  
- Redução significativa de tempo e erros no trabalho administrativo.  

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**  
- **Pandas** (manipulação de dados e geração da planilha)  
- **odfpy** (para exportação em formato ODS compatível com LibreOffice)  
- Outras bibliotecas específicas para leitura dos arquivos de notas (caso necessário).  

---

## 📂 Estrutura do Projeto

A estrutura do código é simples e direta, composta apenas por dois arquivos principais:

```bash
.
├── leitor_pdf.py   # Script principal responsável pela automação
└── README.md       # Documentação do projeto
