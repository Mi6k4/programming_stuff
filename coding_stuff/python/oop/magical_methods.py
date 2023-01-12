class Point:
    MAX_COORD = 100
    MIN_COORD = 0

    def __init__(self,x,y):
        self.x=x
        self.y=y
    def set_coord(self,x,y):
        if self.MIN_COORD <= x <= self.MAX_COORD:

            self.x=x
            self.y=x

    #wrong
    def set_bound(self,left):
        self.MIN_COORD=left # ne pravilno, tak sozdastsa eshe odin atribut eksemplyara classa

    #rigth
    @classmethod
    def set_bound_right(cls,left):
        cls.MIN_COORD=left


    # срабатывает при обращении к любому атрибуту   экземпляра класса
    def __getattribute__(self, item):
        if item == "x":
            raise ValueError("access denied")
        else:
            return object.__getattribute__(self,item)

    # срабатывает при присвоении значения любому атрибуту
    def __setattr__(self, key, value):
        if key=="z":
            raise ValueError("not permitted name")
        else:
            return object.__setattr__(self,key,value)


    #вызывется когда идет обращение к несуществующему атрибуту экземпляра класса
    def __getattr__(self, item):
         return False

    #срабатывает при удалении атрибута класса
    def __delattr__(self, item):
        print("__delattr    __: "+item)
        object.__delattr__(self,item)




pt1 = Point(1,2)
pt2 = Point(10,20)
pt1.set_bound(-100)
print(pt1.__dict__)
print(Point.__dict__) #min_coord ostalsya tem je

a=pt1.y
print(pt1.xx)
del pt1.x
print(pt1.__dict__)

