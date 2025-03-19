#!/usr/bin/env python3
"""
Script para limpar o sistema Moraca, removendo dados de teste e preparando
para uma nova bateria de testes.

Uso:
    python limpar_sistema.py              # Executa a limpeza completa
    python limpar_sistema.py --simular    # Mostra o que seria feito, sem executar
    python limpar_sistema.py --backup     # Cria um backup antes de limpar
    python limpar_sistema.py --modulos    # Remove também módulos não utilizados
    python limpar_sistema.py --ajuda      # Mostra esta mensagem de ajuda
"""

import os
import shutil
import sys
import datetime
import zipfile

def confirmar_acao(mensagem):
    """Solicita confirmação do usuário antes de executar ações destrutivas."""
    resposta = input(f"{mensagem} (s/n): ").lower().strip()
    return resposta == 's'

def criar_estrutura_basica():
    """
    Cria a estrutura básica de diretórios do sistema, caso não exista.
    Esta função garante que, mesmo após a limpeza, a estrutura principal será mantida.
    """
    estrutura_basica = [
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "tecnico", "andamento"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "pendentes"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "aprovados"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "rejeitados"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "convertidos"),
    ]
    
    for diretorio in estrutura_basica:
        if not os.path.exists(diretorio):
            os.makedirs(diretorio)
            print(f"Criado diretório: {diretorio}")

def limpar_arquivos_temporarios():
    """Remove arquivos temporários e diretórios de teste."""
    # Lista de diretórios que podem ser completamente removidos
    diretorios_remover = [
        "temp_estoque",  # Diretório temporário criado pelos testes de estoque
        os.path.join("MORACA", "MORACA1", "ESTOQUE")  # Diretório de estoque (teste)
    ]
    
    # Remover diretórios temporários
    for diretorio in diretorios_remover:
        if os.path.exists(diretorio):
            try:
                shutil.rmtree(diretorio)
                print(f"Removido diretório: {diretorio}")
            except Exception as e:
                print(f"Erro ao remover {diretorio}: {e}")
    
    # Remover arquivos Python temporários
    for raiz, dirs, arquivos in os.walk("."):
        for arquivo in arquivos:
            if arquivo.endswith(".pyc") or arquivo.endswith(".pyo"):
                try:
                    os.remove(os.path.join(raiz, arquivo))
                    print(f"Removido arquivo: {os.path.join(raiz, arquivo)}")
                except Exception as e:
                    print(f"Erro ao remover {os.path.join(raiz, arquivo)}: {e}")

def remover_modulos_nao_utilizados():
    """Remove módulos Python que não serão mais utilizados."""
    modulos_remover = [
        "estoque_manager.py",     # Módulo de gestão de estoque
        "teste_estoque.py",       # Script de teste do estoque
    ]
    
    # Arquivos que podem ser removidos opcionalmente, após confirmação específica
    modulos_opcionais = [
        "orcamentos.py",          # Módulo base de orçamentos (pode ser mantido para referência)
        "teste_orcamentos.py",    # Script de teste de orçamentos
        "integracao_orcamentos.py", # Integração dos orçamentos
        "README_ORCAMENTOS.md",   # Documentação de orçamentos
    ]
    
    # Confirmar remoção de módulos opcionais
    remover_opcionais = confirmar_acao("Deseja remover também os módulos relacionados a orçamentos?")
    
    # Unir as listas se o usuário confirmou
    if remover_opcionais:
        modulos_remover.extend(modulos_opcionais)
    
    # Remover os módulos
    for modulo in modulos_remover:
        if os.path.exists(modulo):
            try:
                os.remove(modulo)
                print(f"Removido módulo: {modulo}")
            except Exception as e:
                print(f"Erro ao remover módulo {modulo}: {e}")

def limpar_documentos_teste():
    """Limpa documentos de teste gerados em execuções anteriores."""
    # Limpar documentos de orçamentos
    base_orcamentos = os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos")
    if os.path.exists(base_orcamentos):
        # Remover controle de orçamentos
        controle_path = os.path.join(base_orcamentos, "controle_orcamentos.xlsx")
        if os.path.exists(controle_path):
            try:
                os.remove(controle_path)
                print(f"Removido arquivo: {controle_path}")
            except Exception as e:
                print(f"Erro ao remover {controle_path}: {e}")
        
        # Limpar orçamentos por status
        for status in ["pendentes", "aprovados", "rejeitados", "convertidos"]:
            status_dir = os.path.join(base_orcamentos, status)
            if os.path.exists(status_dir):
                for arquivo in os.listdir(status_dir):
                    if arquivo.endswith(".xlsx"):
                        try:
                            os.remove(os.path.join(status_dir, arquivo))
                            print(f"Removido orçamento: {os.path.join(status_dir, arquivo)}")
                        except Exception as e:
                            print(f"Erro ao remover {os.path.join(status_dir, arquivo)}: {e}")
    
    # Limpar documentos de OS de andamento
    base_andamento = os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "tecnico", "andamento")
    if os.path.exists(base_andamento):
        # Verificar se deve manter exemplos
        manter_exemplos = confirmar_acao("Deseja manter alguns arquivos de exemplo para referência?")
        exemplos_mantidos = 0
        
        for arquivo in os.listdir(base_andamento):
            if arquivo.endswith(".xlsx"):
                # Opção para manter alguns exemplos para referência
                if manter_exemplos and exemplos_mantidos < 2 and "Ziehm" in arquivo:
                    exemplos_mantidos += 1
                    print(f"Mantido exemplo: {os.path.join(base_andamento, arquivo)}")
                    continue
                
                try:
                    os.remove(os.path.join(base_andamento, arquivo))
                    print(f"Removido documento: {os.path.join(base_andamento, arquivo)}")
                except Exception as e:
                    print(f"Erro ao remover {os.path.join(base_andamento, arquivo)}: {e}")

def criar_backup():
    """
    Cria um backup dos dados antes de limpar.
    
    Returns:
        str: Caminho do arquivo de backup ou None se falhar
    """
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = "backups"
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    
    backup_file = os.path.join(backup_dir, f"moraca_backup_{timestamp}.zip")
    print(f"\n=== Criando backup em {backup_file} ===")
    
    try:
        with zipfile.ZipFile(backup_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Backup dos documentos
            backup_path(zipf, os.path.join("MORACA", "MORACA1", "DOCUMENTOS"))
            
            # Backup de arquivos Python importantes
            for arquivo in os.listdir("."):
                if arquivo.endswith(".py") and os.path.isfile(arquivo):
                    zipf.write(arquivo, arquivo)
                    print(f"Adicionado ao backup: {arquivo}")
        
        print(f"Backup criado com sucesso em: {backup_file}")
        return backup_file
    except Exception as e:
        print(f"Erro ao criar backup: {e}")
        return None

def backup_path(zipf, path):
    """Adiciona um diretório recursivamente ao arquivo zip."""
    if os.path.exists(path):
        if os.path.isdir(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, os.path.dirname(path))
                    try:
                        zipf.write(file_path, arcname)
                        print(f"Adicionado ao backup: {arcname}")
                    except Exception as e:
                        print(f"Erro ao adicionar {file_path} ao backup: {e}")

def limpar_tudo(criar_backup_antes=False, remover_modulos=False):
    """
    Função principal que orquestra o processo de limpeza.
    
    Args:
        criar_backup_antes (bool): Se True, cria um backup antes de limpar
        remover_modulos (bool): Se True, remove módulos não utilizados
    
    Returns:
        bool: True se a limpeza foi bem-sucedida, False caso contrário
    """
    # Verificar se estamos no diretório correto
    if not os.path.exists("MORACA") or not os.path.isdir("MORACA"):
        print("AVISO: O diretório 'MORACA' não foi encontrado no diretório atual.")
        if not confirmar_acao("Deseja continuar mesmo assim?"):
            print("Operação cancelada.")
            return False
    
    # Confirmar a operação
    if not confirmar_acao("Esta operação irá remover todos os arquivos de teste. Continuar?"):
        print("Operação cancelada.")
        return False
    
    # Criar backup se solicitado
    if criar_backup_antes:
        backup_file = criar_backup()
        if not backup_file and not confirmar_acao("Falha ao criar backup. Deseja continuar mesmo assim?"):
            print("Operação cancelada.")
            return False
    
    # Executar limpeza
    print("\n=== Iniciando limpeza do sistema ===")
    
    # 1. Remover arquivos temporários
    print("\n-- Removendo arquivos temporários --")
    limpar_arquivos_temporarios()
    
    # 2. Limpar documentos de teste
    print("\n-- Limpando documentos de teste --")
    limpar_documentos_teste()
    
    # 3. Remover módulos não utilizados, se solicitado
    if remover_modulos:
        print("\n-- Removendo módulos não utilizados --")
        remover_modulos_nao_utilizados()
    
    # 4. Recriar estrutura básica
    print("\n-- Recriando estrutura básica --")
    criar_estrutura_basica()
    
    print("\n=== Limpeza concluída ===")
    return True

def modo_simulacao(verificar_modulos=False):
    """
    Executa o script em modo de simulação, apenas mostrando o que seria feito.
    
    Args:
        verificar_modulos (bool): Se True, inclui informações sobre módulos que seriam removidos
    """
    print("\n=== MODO DE SIMULAÇÃO ===")
    print("As seguintes operações seriam realizadas:")
    
    # Simulação de remoção de diretórios temporários
    print("\nDiretórios que seriam removidos:")
    diretorios = ["temp_estoque", os.path.join("MORACA", "MORACA1", "ESTOQUE")]
    for diretorio in diretorios:
        if os.path.exists(diretorio):
            print(f"- {diretorio}")
    
    # Simulação de limpeza de documentos
    print("\nDiretórios de documentos que seriam limpos:")
    docs_dirs = [
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "pendentes"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "aprovados"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "rejeitados"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "orcamentos", "convertidos"),
        os.path.join("MORACA", "MORACA1", "DOCUMENTOS", "tecnico", "andamento")
    ]
    
    for diretorio in docs_dirs:
        if os.path.exists(diretorio):
            arquivos = [f for f in os.listdir(diretorio) if f.endswith(".xlsx")]
            if arquivos:
                print(f"- {diretorio} ({len(arquivos)} arquivos)")
    
    # Simulação de remoção de módulos
    if verificar_modulos:
        print("\nMódulos que seriam removidos:")
        modulos_remover = ["estoque_manager.py", "teste_estoque.py"]
        for modulo in modulos_remover:
            if os.path.exists(modulo):
                print(f"- {modulo}")
        
        print("\nMódulos que poderiam ser removidos após confirmação:")
        modulos_opcionais = ["orcamentos.py", "teste_orcamentos.py", "integracao_orcamentos.py", "README_ORCAMENTOS.md"]
        for modulo in modulos_opcionais:
            if os.path.exists(modulo):
                print(f"- {modulo}")

    print("\nNenhuma alteração foi realizada. Para executar a limpeza, execute o script sem a opção --simular.")

def mostrar_ajuda():
    """Mostra mensagem de ajuda com opções disponíveis."""
    print(__doc__)
    print("Opções disponíveis:")
    print("  --simular, -s   : Apenas simula o que seria feito, sem executar")
    print("  --backup, -b    : Cria um backup antes de limpar")
    print("  --modulos, -m   : Remove também módulos não utilizados (estoque, etc.)")
    print("  --ajuda, -h     : Mostra esta mensagem de ajuda")
    print("\nExemplos:")
    print("  python limpar_sistema.py                # Executa a limpeza completa")
    print("  python limpar_sistema.py --simular      # Mostra o que seria feito")
    print("  python limpar_sistema.py --backup       # Cria backup antes de limpar")
    print("  python limpar_sistema.py --modulos      # Remove módulos não utilizados")
    print("  python limpar_sistema.py --backup --modulos  # Faz backup e remove módulos não utilizados")

if __name__ == "__main__":
    # Verificar argumentos
    if len(sys.argv) > 1:
        if any(arg in ["--ajuda", "-h"] for arg in sys.argv):
            mostrar_ajuda()
            sys.exit(0)
        
        simular = any(arg in ["--simular", "-s"] for arg in sys.argv)
        backup = any(arg in ["--backup", "-b"] for arg in sys.argv)
        modulos = any(arg in ["--modulos", "-m"] for arg in sys.argv)
        
        if simular:
            modo_simulacao(verificar_modulos=modulos)
            if backup:
                print("\nSe executado com a opção --backup, seria criado um backup antes da limpeza.")
        else:
            limpar_tudo(criar_backup_antes=backup, remover_modulos=modulos)
    else:
        limpar_tudo() 