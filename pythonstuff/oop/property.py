class Person:
    def __init__(self,name,age):
        self.__name = name
        self.__age = age
    @property
    def age(self):
        return self.__age

    @age.setter
    def age(self,age):
        self.__age = age

    @age.deleter
    def old(self):
        del sf.__old

   # age = property()
    #age = age.setter(set_age)
    #age = age.getter(get_age)



p = Person("sergey",20)
#print(p.age)
p.__dict__["age"] = " age "
p.age = 35
#p.age=35
#a = p.age
print(p.age, p.__dict__)