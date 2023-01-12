class Database:

    __instance= None

    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __del__(self):
        Database.__instance= None


    def __init__(self,user,psw,port):
        self.user=user
        self.psw=psw
        self.port=port

    def connection(self):
        print(f"connection with DB: {self.user},{self.psw},{self.port}")

    def close(self):
        print("close connection with DB")
    def read(self):
        return "database data"

    def write(self,data):
        print(f"database record {data}")


db = Database('root','1234',80)
db2= Database('root2','12345',40)

db.connection()
db2.connection()


print(id(db),id(db2))