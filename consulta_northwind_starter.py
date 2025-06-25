
import pyodbc

class BaseDadosAccess:
    def __init__(self, caminho_ficheiro):
        self.caminho = caminho_ficheiro
        self.conn = None
        self.cursor = None
        self._ligar()

    def _ligar(self):
        try:
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

class Produto:
    def __init__(self, nome, preco):
        self.nome = nome
        self.preco = preco

    def __str__(self):
        return f"{self.nome} - €{self.preco:.2f}"

# Código principal
if __name__ == "__main__":
    # Substituir pelo caminho real da vossa base de dados
    caminho_bd = r"Database1.accdb"

    bd = BaseDadosAccess(caminho_bd)

    sql = """
    SELECT ProductName, UnitPrice
    FROM Products
    ORDER BY UnitPrice DESC;
    """

    resultados = bd.executar_consulta(sql)

    produtos = [
        Produto(nome=row.ProductName, preco=row.UnitPrice)
        for row in resultados
    ]

    print("📋 Produtos ordenados por preço (descendente):\n")
    for i, produto in enumerate(produtos[:10], start=1):
        print(f"{i}. {produto}")

    bd.fechar()
