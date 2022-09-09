import cx_Oracle

def importOracle():
    cx_Oracle.init_oracle_client(lib_dir=r"/Applications/instantclient_19_8")
    connection = cx_Oracle.connect("sjj/123@10.211.55.3:1521/orcl")
    cursor = connection.cursor()
    cursor.execute("select 1 from dual")
    data = cursor.fetchone()
    print(data)

if __name__ == '__main__':
    importOracle()
