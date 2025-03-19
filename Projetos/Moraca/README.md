# Moraca - Sistema de Ordem de Serviço

Sistema simples para gerenciamento de Ordens de Serviço (OS) desenvolvido com Python e Tkinter.

## Requisitos

- Python 3.7 ou superior
- Dependências listadas em `requirements.txt`

## Instalação

1. Clone ou baixe este repositório
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

## Como Usar

Execute o arquivo principal:

```bash
python main.py
```

### Funcionalidades

- Criar nova Ordem de Serviço
- Consultar OS existentes
- Pesquisar OS por número ou cliente
- Visualizar últimas OS criadas
- Gerenciar status das OS

### Estrutura do Banco de Dados

O sistema utiliza SQLite para armazenar as OSs com os seguintes campos:

- Número da OS (formato: OSYYXXX)
- Nome do Cliente
- Tipo de Máquina
- Descrição do Problema
- Nível de Urgência
- Data
- Status
