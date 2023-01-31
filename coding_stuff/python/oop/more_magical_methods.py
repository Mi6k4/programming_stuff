#dunder methods == double underscore
class Cat:
    def __init__(self,name):
        self.name=name
#служит для вывода отладочной инфы
    def __repr__(self):
        return f"{self.__class__}: {self.name}"
#служит для вывода информации для пользователя
    def __str__(self):
        return f"{self.name}"

class Point:
    def __init__(self,*args):
        self.__coords = args
    def __len__(self):
        return len(self.__coords)
    def __abs__(self):
        return list(map(abs,self.__coords))

cat = Cat("Vasya")
print(cat)

p = Point(1,2)
print(len(p))
print(abs((p)))



