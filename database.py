import sqlite3

class Data_base:

    def __init__(self, name = 'system.db') -> None:        
        self.name = name

    def connect(self):
        self.connection = sqlite3.connect(self.name)

    def close_connection(self):
        try:
            self.connection.close()
        except:
            pass
    
    def create_table(self):
        cursor = self.connection.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS visitantes(id INTEGER PRIMARY KEY AUTOINCREMENT, cpf TEXT NOT NULL UNIQUE, nome TEXT, orgao TEXT, municipio TEXT, telefone TEXT, email TEXT);           
        """)
    
    def insert_table(self, fullDataSet):

        campos_tabela = ('cpf','nome','orgao','municipio','telefone','email')

        qntd = ("?,?,?,?,?,?")
        cursor = self.connection.cursor()

        try:
            cursor.execute(f"""INSERT INTO visitantes{campos_tabela}
            VALUES({qntd})""", fullDataSet)
            self.connection.commit()
            return("OK")

        except:
            return "Erro"
    
    def create_table_2(self):
        cursor = self.connection.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS registros(id INTEGER PRIMARY KEY AUTOINCREMENT, cpf TEXT, nome TEXT, orgao TEXT, municipio TEXT, telefone TEXT, email TEXT, data TEXT, destino TEXT);           
        """)
    
    def insert_table_2(self, fullDataSet):

        campos_tabela = ('cpf','nome','orgao','municipio','telefone','email','data','destino')

        qntd = ("?,?,?,?,?,?,?,?")
        cursor = self.connection.cursor()

        try:
            cursor.execute(f"""INSERT INTO registros{campos_tabela}
            VALUES({qntd})""", fullDataSet)
            self.connection.commit()
            return("OK")

        except:
            return "Erro"

    def select_all(self):
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros")
            registros = cursor.fetchall()
            return registros
        except:
            pass

    def select_nomes(self):
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM visitantes")
            registros = cursor.fetchall()
            return registros
        except:
            pass
     
