import sys
import cx_Oracle

cx_Oracle.init_oracle_client(lib_dir="/Applications/instantclient_19_8")

if __name__ == '__main__':
    print(f'系统信息版本: {sys.version}')

    # 操作数据库
    db = cx_Oracle.connect('sjj', '123', 'localhost:1521/helowin')
    with db.cursor() as cur:
        for row in cur.execute("select * from cz_gk_yszbb where rownum <= 10"):
            print(row)
    cur.close()
    db.close()

