from string import ascii_letters

class Person:
    S_RUS = "абвгдеёжзийклмнопрстуфхцчшщьыъэюя-"
    S_RUS_UPPER=S_RUS.upper()
    def __init__(self,fio,old,ps,weight):
        self.verify_fio(fio)
        self.verify_old(old)
        self.verify_ps(ps)
        self.verify_weight(weight)
        self.__fio= fio.split()
        self.__old= old
        self.__passport= ps
        self.__weight= weight

    @classmethod
    def verify_fio(cls, fio):
        if type(fio) != str:
            raise TypeError("FIO must be a str")

        f = fio.split()
        if len(f) != 3:
            raise TypeError("incorrect format")

        letters = ascii_letters + cls.S_RUS + cls.S_RUS_UPPER

        for s in f:
            if len(s)<1:
                raise TypeError("must be at least 1 letter")
            if len(s.strip(letters)) != 0:
                raise TypeError("must be letters")

    @classmethod
    def verify_old(cls,old):
        if type(old) != int or  old < 14 or old >120:
            raise TypeError("must be int and number")

    @classmethod
    def verify_weight(cls, w):
        if type(w) != float or w < 20 :
            raise TypeError("must be float and >20")

    @classmethod
    def verify_ps(cls,ps):
        if type(ps)!=str:
            raise TypeError("ps must be str")
        s = ps.split()
        if len(s) != 2 or len(s[0]) != 4 or len(s[1]) != 6:
            raise  TypeError("incorrect format")
        for p in s:
            if not p.isdigit():
                raise TypeError("must be digits")

    @property
    def fio(self):
        return self.__fio

    @property
    def old(self):
        return self.__old
    @old.setter
    def old(self,old):
        self.verify_old(old)
        self.__old=old

    @property
    def weight(self):
        return self.__weight

    @weight.setter
    def weight(self, weight):
        self.verify_weight(weight )
        self.__weight = weight

    @property
    def ps(self):
        return self.__ps

    @ps.setter
    def ps(self, ps):
        self.verify_ps(ps)
        self.__ps = ps


p = Person("qweqr aret asf", 30, "1231 134235", 80.0)

p.old = 100
p.passport = "1234 567890"
p.weight = 70.0

print(p.__dict__)

