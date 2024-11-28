import mysql.connector
from faker import Faker
import random
from unidecode import unidecode

fake = Faker('pt_BR')

# Conexão ao MySQL
mydb = mysql.connector.connect(
    host="localhost",
    user="root",  
    password=""  
)

mycursor = mydb.cursor()

mycursor.execute("USE livraria_109")

mycursor.execute("""
CREATE TABLE IF NOT EXISTS livros (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome_do_livro VARCHAR(50),
    nome_da_editora VARCHAR(50),
    uf_editora VARCHAR(2),
    ano_publicacao INT,
    num_paginas INT,
    nome_autor VARCHAR(50),
    sexo_autor CHAR(1),
    cpf_autor VARCHAR(11)
)
""")

mydb.commit()
print("Banco de dados e tabela criados com sucesso.")

# Função para limpar os nomes (remover acentos e títulos)
def clean_name(name):
    name = name.replace("Sr. ", "").replace("Sra. ", "")
    return unidecode(name)

def cpf_validar():
    cpf = fake.random_number(digits=9, fix_len=True)
    cpf_str = f"{cpf:09d}"
    return cpf_str

def insert_livro():
    nome_do_livro = fake.catch_phrase()
    nome_da_editora = fake.company()
    uf_editora = random.choice(['SP', 'RJ', 'DF'])
    ano_publicacao = random.randint(2000, 2024)
    num_paginas = random.randint(50, 500)
    sexo_autor = random.choice(['M', 'F'])
    nome_autor = fake.name_male() if sexo_autor == 'M' else fake.name_female()
    nome_autor = clean_name(nome_autor)  # Limpa o nome do autor
    nome_do_livro = clean_name(nome_do_livro)  # Limpa o nome do livro
    nome_da_editora = clean_name(nome_da_editora)  # Limpa o nome da editora
    cpf_autor = cpf_validar()

    sql = """
    INSERT INTO livros (nome_do_livro, nome_da_editora, uf_editora, ano_publicacao, num_paginas, nome_autor, sexo_autor, cpf_autor)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """
    val = (nome_do_livro, nome_da_editora, uf_editora, ano_publicacao, num_paginas, nome_autor, sexo_autor, cpf_autor)
    mycursor.execute(sql, val)

# Gerar 50 registros de livros
for _ in range(50):
    insert_livro()

mydb.commit()
print("50 registros inseridos com sucesso.")

# Fechar a conexão
mycursor.close()
mydb.close()
