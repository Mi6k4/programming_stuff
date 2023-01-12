class ThreadData:
    __shared_stats={
        "name":'thread1',
        "data":{},
        "id":1
    }

    def __init__(self):
        self.__dict__ = self.__shared_stats



th1 = ThreadData()

th2 = ThreadData()

print(th1.__dict__)
print(th2.__dict__)

th2.id=3

print(th1.__dict__)
print(th2.__dict__)

th1.attrnew= "new attr"

print(th1.__dict__)
print(th2.__dict__)
