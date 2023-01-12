# attribute - public - available everywhere
# _attribute - protected - available only in the class and child classes
# __attribute - private - only in the class

class Point:
    def __init__(self,x=0,y=0):
        self.__x=0
        self.__y=0
        if self.__check_value(x) and self.__check_value(y):
            self.__x=x
            self.__y=y
    @classmethod
    def __check_value(cls,x):
        return type(x) in (int,float)


    def setcoord(self,x,y):
        if self.__check_value(x)  and self.__check_value(y):
            self.__x = x
            self.__y = y
        else:
            raise ValueError("coordinates must be numbers")
    def getcoord(self):
        return self.__x,self.__y


pt = Point(1,2)

pt.setcoord(10,20)
print(dir(pt))

print(pt._Point__x) # you can find this with dir(pt)

