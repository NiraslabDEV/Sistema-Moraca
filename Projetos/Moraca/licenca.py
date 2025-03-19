import os
import hashlib
import datetime
import json
import base64
from datetime import datetime, timedelta

# Chave secreta para validação de licença (em ambiente real, use uma mais complexa e armazene com segurança)
CHAVE_SECRETA = "8Fd7GbP3sQl9WnX4eR5tY6uI7oA2zJ0m"
ARQUIVO_LICENCA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "licenca.dat")

def gerar_codigo_licenca(cliente, data_expiracao=None, tipo_licenca="Comercial"):
    """
    Gera um código de licença baseado na data de expiração.
    
    Args:
        cliente (str): Nome do cliente ou empresa.
        data_expiracao (str): Data de expiração no formato "YYYY-MM-DD".
                         Se não fornecido, usa um ano a partir da data atual.
        tipo_licenca (str): Tipo de licença: "Comercial", "Demo", "Educacional".
    
    Returns:
        str: Código de licença codificado em base64.
    """
    # Se não houver data de expiração, definir para um ano após a data atual
    if not data_expiracao:
        data_atual = datetime.now()
        data_expiracao = (data_atual + timedelta(days=365)).strftime("%Y-%m-%d")
    
    # Criar dicionário com informações da licença
    info_licenca = {
        "cliente": cliente,
        "data_expiracao": data_expiracao,
        "data_geracao": datetime.now().strftime("%Y-%m-%d"),
        "tipo_licenca": tipo_licenca
    }
    
    # Converter para JSON
    json_licenca = json.dumps(info_licenca, sort_keys=True)
    
    # Criar assinatura
    assinatura = hashlib.sha256((json_licenca + CHAVE_SECRETA).encode()).hexdigest()
    
    # Adicionar assinatura ao dicionário
    info_licenca["assinatura"] = assinatura
    
    # Codificar em base64
    codigo_final = base64.b64encode(json.dumps(info_licenca).encode()).decode()
    
    return codigo_final

def verificar_licenca(codigo_licenca):
    """
    Verifica se um código de licença é válido.
    
    Args:
        codigo_licenca (str): Código de licença a ser verificado.
    
    Returns:
        dict: Dicionário com status da licença (valida, mensagem, etc.)
    """
    resultado = {
        "valida": False,
        "mensagem": "Código de licença inválido.",
        "cliente": "",
        "tipo_licenca": "",
        "data_expiracao": "",
        "dias_restantes": 0
    }
    
    try:
        # Decodificar o código
        json_licenca = base64.b64decode(codigo_licenca.encode()).decode()
        info_licenca = json.loads(json_licenca)
        
        # Extrair informações
        cliente = info_licenca.get("cliente", "")
        data_expiracao = info_licenca.get("data_expiracao", "")
        tipo_licenca = info_licenca.get("tipo_licenca", "")
        assinatura = info_licenca.pop("assinatura", "")
        
        # Recalcular assinatura para verificar
        json_sem_assinatura = json.dumps(info_licenca, sort_keys=True)
        assinatura_calculada = hashlib.sha256((json_sem_assinatura + CHAVE_SECRETA).encode()).hexdigest()
        
        # Verificar se a assinatura é válida
        if assinatura != assinatura_calculada:
            resultado["mensagem"] = "Assinatura de licença inválida."
            return resultado
        
        # Verificar data de expiração
        data_atual = datetime.now()
        data_exp = datetime.strptime(data_expiracao, "%Y-%m-%d")
        
        dias_restantes = (data_exp - data_atual).days
        
        # Atualizar resultados
        resultado["cliente"] = cliente
        resultado["tipo_licenca"] = tipo_licenca
        resultado["data_expiracao"] = data_expiracao
        resultado["dias_restantes"] = dias_restantes
        
        # Verificar se a licença expirou
        if dias_restantes < 0:
            resultado["mensagem"] = f"Licença expirada em {data_expiracao}."
            return resultado
        
        # Licença válida
        resultado["valida"] = True
        resultado["mensagem"] = f"Licença válida até {data_expiracao} ({dias_restantes} dias restantes)."
        
        return resultado
        
    except Exception as e:
        resultado["mensagem"] = f"Erro ao verificar licença: {str(e)}"
        return resultado

def salvar_licenca(codigo_licenca):
    """
    Salva o código de licença em um arquivo.
    
    Args:
        codigo_licenca (str): Código de licença a ser salvo.
    
    Returns:
        bool: True se salvou com sucesso, False caso contrário.
    """
    try:
        with open(ARQUIVO_LICENCA, "w") as f:
            f.write(codigo_licenca)
        return True
    except Exception as e:
        print(f"Erro ao salvar licença: {str(e)}")
        return False

def carregar_licenca():
    """
    Carrega o código de licença do arquivo.
    
    Returns:
        str: Código de licença ou string vazia se não encontrado.
    """
    try:
        if os.path.exists(ARQUIVO_LICENCA):
            with open(ARQUIVO_LICENCA, "r") as f:
                return f.read().strip()
        return ""
    except Exception as e:
        print(f"Erro ao carregar licença: {str(e)}")
        return ""

def gerar_nova_licenca_demo(dias=30):
    """
    Gera uma nova licença de demonstração válida por um período específico.
    
    Args:
        dias (int): Número de dias de validade da licença demo.
    
    Returns:
        str: Código de licença demo.
    """
    data_atual = datetime.now()
    data_expiracao = (data_atual + timedelta(days=dias)).strftime("%Y-%m-%d")
    
    return gerar_codigo_licenca(
        cliente="Cliente Demo", 
        data_expiracao=data_expiracao, 
        tipo_licenca="Demo"
    )

# Demonstração de uso
if __name__ == "__main__":
    # Gerar licença para um ano
    codigo = gerar_codigo_licenca("MORACA Sistemas", tipo_licenca="Comercial")
    print(f"Código de licença gerado: {codigo}")
    
    # Verificar a licença
    resultado = verificar_licenca(codigo)
    print(f"Verificação: {resultado['mensagem']}")
    
    # Salvar a licença
    if salvar_licenca(codigo):
        print("Licença salva com sucesso!")
    
    # Carregar a licença
    codigo_carregado = carregar_licenca()
    if codigo_carregado:
        print("Licença carregada com sucesso!") 