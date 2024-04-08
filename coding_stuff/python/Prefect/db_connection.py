from prefect import task, flow
import psycopg2

credentials_gp = {'host': '', 'database': 'warehouse', 'port': 6432,
                      'user': 'zeppelin',
                      'password': '', 'schema': 'public'}

@task
def update_task():
    print('connecting')
    conn_gp = psycopg2.connect(
        host=credentials_gp['host'],
        database=credentials_gp['database'],
        port=credentials_gp['port'],
        user=credentials_gp['user'],
        password=credentials_gp['password']
    )
    print('connected')
    cursor = conn_gp.cursor()
    cursor.execute('select 1')
    answer=cursor.fetchall()[0]
    return answer

@flow(log_prints=True)
def db_flow():
    task_result=update_task
    print(task_result)

db_flow()