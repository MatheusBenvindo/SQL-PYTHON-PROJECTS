import mysql.connector
import random

# Conexão ao MySQL para acessar o banco de dados
mydb = mysql.connector.connect(
    host="localhost",
    user="root",  # Usuário root do MariaDB
    password="",  
    database="livraria_109"  #BD
)

mycursor = mydb.cursor()

# Função para atualizar os livros com preco = 0
def update_livros_com_preco_zero():
    novo_preco = round(random.uniform(20.01, 200.00), 2)  # Gera um novo preço entre 20.01 e 200.00
    sql = "UPDATE livros SET preco = %s WHERE preco = 0"
    mycursor.execute(sql, (novo_preco,))

# Atualizar livros com preco = 0
update_livros_com_preco_zero()

# Confirmar as alterações no banco de dados
mydb.commit()

print("Preços dos livros com preço 0 atualizados.")

# Fechar a conexão
mycursor.close()
mydb.close()
