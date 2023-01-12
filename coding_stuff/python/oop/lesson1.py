class Point:
    "class fo point coordinates"
    color = 'red'
    circle = 2



a = Point()
b = Point()


Point.circle = 1

a.color='green'

print(Point.color)
print(Point.circle)
print(Point.__dict__)
print(type(a))
print(a.__dict__)

Point.type_pt = 'disc'
print(Point.__dict__)

setattr(Point,'prop',1)
print(Point.__dict__)
setattr(Point,'type_pt','square')
print(Point.__dict__)

print(getattr(Point, 'a',False))
print(getattr(Point,'color',False))
print(getattr(Point,'color'))


del Point.prop
print(hasattr(Point,'prop'))

delattr(Point,"type_pt")
print(hasattr(Point,"color"))
print(hasattr(Point,"type_pt"))
print(hasattr(a,"color"))

a.x=1
a.y=2

b.x=10
b.y=20

print(Point.__doc__)