class Point:
    color = 'red'
    circle = 2

    def __init__(self,x,y):
        print("init call")
        self.x=x
        self.y=y

    def __del__(self):
        print("Deleting exemplar:"+str(self))

    def set_coords(self,x,y):
        print("set_coords method call" + str(self))
        self.x = x
        self.y = y
    def get_coords(self):
        return (self.x,self.y)



pt = Point(1,2)
print(pt.__dict__)