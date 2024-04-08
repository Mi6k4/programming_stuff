from prefect import task,flow
import psycopg2


@task
def create_message():
    msg = 'Hello from task'
    return msg

@flow(log_prints=True)
def hello():
    task_message=create_message()
    print(task_message)


hello()