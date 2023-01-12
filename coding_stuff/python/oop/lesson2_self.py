class Point:
    color = 'red'
    circle = 2

    def set_coords(self,x,y):
        print("set_coords method call" + str(self))
        self.x = x
        self.y = y
    def get_coords(self):
        return (self.x,self.y)




#print(Point.set_coords)
#Point.set_coords()

pt = Point()

print(pt.set_coords)
pt.set_coords(1,2)
print(pt.__dict__)
#Point.set_coords(pt)

pt2=Point()
pt2.set_coords(10,20)
print(pt2.__dict__)
print(pt.get_coords())

f = getattr(pt,"get_coords")
print(f)
print(f())
