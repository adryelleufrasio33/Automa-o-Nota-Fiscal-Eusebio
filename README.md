# Automação de Extração de Dados de Nota Fiscal – Prefeitura de Eusébio

Este projeto é uma **automação em Python** que lê **notas fiscais em PDF** emitidas pela Prefeitura de Eusébio (CE), extrai informações relevantes e gera uma planilha organizada com os dados.  

O objetivo é **eliminar o trabalho manual do setor administrativo**, reduzindo erros e economizando tempo, já que antes era necessário preencher planilhas manualmente com os valores das notas.

---

## 🚀 Funcionalidades

- Leitura automática de arquivos PDF em uma pasta definida no código.  
- Extração de informações-chave, como:  
  - Número da Nota Fiscal  
  - Razão Social  
  - Descrição dos Serviços  
  - Valor dos Serviços  
  - ISS (Valor)  
  - ISS Retido  
- Geração de uma **planilha Excel (`.xlsx`)** com os dados organizados.  
- Abertura automática do arquivo gerado no **LibreOffice Calc** (ou Excel, se disponível).  

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**  
- **pdfplumber** (extração de texto e tabelas de PDFs)  
- **re (regex)** (identificação de padrões nos textos)  
- **openpyxl** (criação da planilha Excel)  
- **subprocess / os** (para abrir o arquivo no LibreOffice ou Excel automaticamente)  

---

## 📂 Estrutura do Projeto

```bash
.
├── leitor_pdf.py   # Script principal responsável pela automação
└── README.md       # Documentação do projeto
