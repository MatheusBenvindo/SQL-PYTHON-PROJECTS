import mysql.connector
from faker import Faker
import random
from unidecode import unidecode

#Faker para gerar dados brasileiros
fake = Faker('pt_BR')

# Conexão ao MySQL para criar o banco de dados e a tabela
mydb = mysql.connector.connect(
    host="localhost",
    user="root",  #root do MariaDB
    password="",  
    database="livraria_109"  # Banco de dados
)

mycursor = mydb.cursor()

# Adicionar a coluna "preco" à tabela "livros" (se ainda não existir)
mycursor.execute("ALTER TABLE livros ADD COLUMN IF NOT EXISTS preco DECIMAL(10, 2) NOT NULL")

# Função para limpar os nomes (remover acentos e títulos para evitar conflitos e formatações dentro do bd)
def clean_name(name):
    name = name.replace("Sr. ", "").replace("Sra. ", "")
    return unidecode(name)

def update_preco_livro():
    # Garante que o preço está dentro da faixa desejada e diferente de 0
    preco = round(random.uniform(25, 251), 2)
    if preco == 0:
        preco = round(random.uniform(20.01, 200.00), 2)
    sql = "UPDATE livros SET preco = %s WHERE id = %s"
    # Seleciona um livro aleatório para atualizar o preço
    livro_id = random.randint(1, 50)  # Supõe que há pelo menos 50 livros
    mycursor.execute(sql, (preco, livro_id))

for _ in range(50):
    update_preco_livro()

mydb.commit()
print("Preços atualizados com sucesso.")

# Fechar a conexão
mycursor.close()
mydb.close()
