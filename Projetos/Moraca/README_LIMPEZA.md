# Script de Limpeza do Sistema Moraca

Este script foi desenvolvido para preparar o sistema Moraca para novos testes, removendo dados e arquivos temporários, mantendo apenas a estrutura básica necessária.

## Funcionalidades

O script de limpeza oferece as seguintes funcionalidades:

1. **Limpeza de Documentos**: Remove arquivos Excel de teste gerados anteriormente.
2. **Remoção de Diretórios Temporários**: Elimina pastas temporárias criadas durante testes.
3. **Backup Automático**: Cria um backup de todos os dados antes de realizar a limpeza.
4. **Remoção de Módulos**: Opção para remover módulos Python que não serão mais utilizados.
5. **Preservação de Exemplos**: Possibilidade de manter alguns arquivos de exemplo para referência.
6. **Modo de Simulação**: Visualização prévia de todas as alterações que serão realizadas.

## Como Usar

### Opções Disponíveis

```
python limpar_sistema.py              # Executa a limpeza completa
python limpar_sistema.py --simular    # Mostra o que seria feito, sem executar
python limpar_sistema.py --backup     # Cria um backup antes de limpar
python limpar_sistema.py --modulos    # Remove também módulos não utilizados
python limpar_sistema.py --ajuda      # Mostra mensagem de ajuda
```

Você pode combinar diferentes opções:

```
python limpar_sistema.py --backup --modulos   # Cria backup e remove módulos
python limpar_sistema.py --simular --modulos  # Simula remoção de módulos
```

### Fluxo Recomendado

Para uma limpeza segura do sistema, recomendamos seguir este fluxo:

1. **Simulação**: Execute primeiro com a opção `--simular` para ver o que será afetado.

   ```
   ./limpar_sistema.py --simular
   ```

2. **Backup**: Antes de executar a limpeza real, crie um backup.

   ```
   ./limpar_sistema.py --backup
   ```

3. **Limpeza Completa**: Se desejar remover também módulos não utilizados, adicione a opção `--modulos`.
   ```
   ./limpar_sistema.py --backup --modulos
   ```

## O que é Removido

### Diretórios

- `temp_estoque/`: Diretório temporário criado pelos testes de estoque
- `MORACA/MORACA1/ESTOQUE/`: Diretório de estoque (teste)

### Documentos

- `MORACA/MORACA1/DOCUMENTOS/orcamentos/`: Orçamentos de teste
- `MORACA/MORACA1/DOCUMENTOS/tecnico/andamento/`: Ordens de serviço de teste

### Módulos (com a opção `--modulos`)

- `estoque_manager.py`: Módulo de gestão de estoque
- `teste_estoque.py`: Script de teste do estoque

Módulos opcionais (removidos após confirmação adicional):

- `orcamentos.py`: Módulo base de orçamentos
- `teste_orcamentos.py`: Script de teste de orçamentos
- `integracao_orcamentos.py`: Integração de orçamentos
- `README_ORCAMENTOS.md`: Documentação de orçamentos

## Backups

Quando executado com a opção `--backup`, o script cria um arquivo ZIP na pasta `backups/` com o seguinte padrão de nome:

```
backups/moraca_backup_AAAAMMDD_HHMMSS.zip
```

Este arquivo contém:

- Todos os documentos da pasta `MORACA/MORACA1/DOCUMENTOS/`
- Todos os arquivos Python do diretório principal

## Precauções

- **Execute sempre em modo de simulação primeiro** para verificar o que será afetado.
- Faça um backup antes de realizar a limpeza completa.
- O script solicita confirmação antes de executar operações destrutivas.
- Você pode optar por manter alguns arquivos de exemplo para referência.

## Restauração

Se necessário restaurar um backup:

1. Localize o arquivo de backup na pasta `backups/`
2. Extraia o conteúdo para o diretório principal do projeto
3. Isso irá restaurar todos os documentos e arquivos Python que estavam presentes no momento do backup
