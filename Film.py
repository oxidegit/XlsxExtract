#电影实体类
class Film:
    def __init__(self, filmInfo, month):
        self.date = filmInfo[0]
        self.name = filmInfo[1]
        self.country = filmInfo[2]
        self.company = filmInfo[3]#出品方
        self.money = [0 for i in range(12)]
        self.count = filmInfo[4] #场次
        self.peopleNum = filmInfo[5]
        self.money[month] = filmInfo[6]
        self.lastMoney = filmInfo[7]
        self.director = filmInfo[8]
        self.writer = filmInfo[9]
        self.actor = filmInfo[10]
        self.province = None
        self.city = None
        self.totalMoney = filmInfo[6]
