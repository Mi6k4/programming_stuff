class Vector:

    MIN_COORD = 0
    MAX_COORD= 100

    @classmethod
    def validate(cls,arg):
        return cls.MIN_COORD<=arg <= cls.MAX_COORD

    def __init__(self,x,y):
        self.x=0
        self.y=0
        if self.validate(x) and self.validate(y):
       # if Vector.validate(x) and Vector.validate(y):
            self.x=x
            self.y=y
        print(self.norm2(self.x, self.y))

    def get_coord(self):
        return self.x,self.y


    @staticmethod
    def norm2(x,y):
        return x*x + y*y # + Vector.MAX_COORD  not reccomended



v = Vector(1,2)
print(Vector.validate(5))
print(Vector.norm2(5,6))
res = v.get_coord()
print(res)