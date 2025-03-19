import sqlite3

conn = sqlite3.connect('moraca.db')
cur = conn.cursor()

# Verificar o esquema da tabela
print("Esquema da tabela 'os':")
cur.execute('PRAGMA table_info(os)')
for col in cur.fetchall():
    print(col)

# Verificar se existem dados
print("\nPrimeiras linhas da tabela 'os':")
try:
    cur.execute('SELECT * FROM os LIMIT 3')
    rows = cur.fetchall()
    if rows:
        for row in rows:
            print(row)
    else:
        print("Nenhum dado encontrado.")
except Exception as e:
    print(f"Erro ao consultar dados: {e}")

conn.close() 