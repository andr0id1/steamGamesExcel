import requests


class GameInfo():

    def __init__(self, countries):
        self.countries = countries

    def getGameInfo(self, id):
        dic = {}
        for country in self.countries:
            resp = requests.get('http://store.steampowered.com/api/appdetails/?appids=' + id + '&cc=' + country)
            datastore = resp.json()
            dic["name"] = datastore[id]["data"]["name"]
            dic[country] = datastore[id]["data"]["price_overview"]["final"]
        return dic
