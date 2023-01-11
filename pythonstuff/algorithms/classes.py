class House():
    def __init__(self,street,number):
        self.street=street
        self.number=number
        self.age=0

    def build(self):
        print("dom na ulitse " + self.street + "pod monerom " + str(self.number) + "postroen")
    def age_of_house(self,year):
        self.age=year+self.age



class Prospekt_house(House):
    def __init__(self, prospect, number):
        super().__init__(self, number)
        self.prospect=prospect

House1= House("Moskovskaya",20)
House2=House("Moskovskaya",21)


print(House1.street)
House2.build()
print(House1.age)
House1.age_of_house(5)
print(House1.age)

PrHouse=Prospekt_house("Lenina",5)
print(PrHouse.prospect)