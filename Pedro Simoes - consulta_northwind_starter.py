# Consulta à Base de Dados Northwind Starter - Pedro Simões

import pyodbc
import os

# Classe responsável pela ligação à base de dados Access
class BaseDadosAccess:
    def __init__(self, caminho_ficheiro):
        self.caminho = caminho_ficheiro
        self.conn = None
        self.cursor = None
        self._ligar()

    def _ligar(self):
        try:
            # Ligação à base de dados Access (.accdb)
            self.conn = pyodbc.connect(
                rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.caminho};"
            )
            self.cursor = self.conn.cursor()
        except Exception as e:
            print("Erro ao ligar à base de dados:", e)

    def executar_consulta(self, sql):
        try:
            self.cursor.execute(sql)
            return self.cursor.fetchall()
        except Exception as e:
            print("Erro ao executar a consulta:", e)
            return []

    def fechar(self):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

# Classe que representa uma encomenda (Order) com os dados do cliente
class Encomenda:
    def __init__(self, id, data, cliente, cidade):
        self.id = id  # OrderID
        self.data = data  # OrderDate
        self.cliente = cliente  # CustomerName
        self.cidade = cidade  # City

    def __str__(self):
        return f"Encomenda {self.id} - {self.data:%Y-%m-%d} - {self.cliente} ({self.cidade})"

# Código principal
if __name__ == "__main__":
    # Caminho para a base de dados Access
    caminho_bd = r".\Database1.accdb"

    # Criar instância da base de dados
    bd = BaseDadosAccess(caminho_bd)

    # Consulta SQL que junta Orders com Customers, ordenada por data descrescente e limita a 20 resultados
    sql = """
    SELECT TOP 20 Orders.OrderID, Orders.OrderDate, Customers.CustomerName, Customers.City
    FROM Orders
    INNER JOIN Customers ON Orders.CustomerID = Customers.CustomerID
    ORDER BY Orders.OrderDate DESC;
    """

    # Executar a consulta e guardar os resultados
    resultados = bd.executar_consulta(sql)

    # Criar lista de objetos Encomenda a partir dos resultados
    encomendas = [
        Encomenda(id=row.OrderID, data=row.OrderDate, cliente=row.CustomerName, cidade=row.City)
        for row in resultados
    ]

    if os.name == 'nt':
        os.system("cls")
    else:
        os.system("clear")

    # Apresentar as 20 encomendas mais recentes
    print("📋 Últimas 20 Encomendas:\n")
    for i, encomenda in enumerate(encomendas, start=1):
        print(f"{i}. {encomenda}")

    # Fechar a ligação à base de dados
    bd.fechar()
