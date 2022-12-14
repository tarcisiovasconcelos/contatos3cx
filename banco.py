import pyodbc

# Conexão com o banco
server = '' 
database = '' 
username = '' 
password = '' 
cnxn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};'
						'SERVER='+server+';'
						'DATABASE='+database+';'
						'UID='+username+';'
						'PWD='+ password)
cursor = cnxn.cursor()

#QUERYS
sqlInsert = 'INSERT INTO tabela (id,setor,telefone) VALUES (?,?,?)'
sqlSelect = 'SELECT setor,telefone FROM tabela WHERE id = ?'
sqlUpdateSetor = 'UPDATE tabela SET setor = ? WHERE id = ?'
sqlUpdateTelefone = 'UPDATE tabela SET telefone = ? WHERE id = ?'
sqlSelectALL = 'SELECT * FROM tabela'
sqlDelete = 'DELETE FROM TABELA WHERE id=?'