
import sqlite3
from XlsxZadania import XlsxZadania

path = 'test.db'
conn = sqlite3.connect(path)
cursor = conn.cursor()
sql = "SELECT * FROM CLIENTS LIMIT 100;"
cursor.execute(sql)
fa = cursor.fetchall()

xlsx = XlsxZadania.nowy("nowy.xlsx", trybtablicaStr=True)
xlsx.dodajArkusz("nazwa", "sql")
xlsx.listaArkuszyUkrytych |= {"sql"}

xlsx.zapisz(fa, [t1[0] for t1 in cursor.description])
xlsx.zapisz([[sql]])
xlsx.zamknij()
