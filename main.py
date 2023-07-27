from collections import defaultdict

import gspread
import pandas as pd
import requests
from xlsxwriter import Workbook

credentials = {
  "type": "service_account",
  "project_id": "bcp-21685",
  "private_key_id": "87a06c94ddb1d5cee2250170e6958111735d7a8d",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC+BHKXOHf8yrJL\nucVEjp+Ff9MhTk1XaJ7pwchBASAQ1ibrAslnaWHnXRwxXRTKPiF1wYblVkxyORgL\nrl7IF41NCMaX0i2NAbL5FnufIlAfOSyQbysv7R5vTGR116U1mm6QX4wD0xf+dfkS\nB8rCT2c2RddSMqlVD+LY9GxZh46xsVeSJX1ryAiAaaU4m2ohSwZRJY9pCJUV9YgN\nmTLyzmHxYjvW6OfYE4iiEubYVaeG1FOE/diqjgAdq1VUvxGF1gq78XV+zAD3WdoF\n1pa7A+U/WfgL1vBy5FBxvhgJnNXq9MzSrIr0MWv1t9ZBM4lFUrJpk8S93w/c3R1T\nQe11gmFlAgMBAAECggEAKMcbEVVCuKsGKsaV7UG/PZ7CIChl5B91VPxrXXngmARh\nKzd4odrlcgmTCovrcYB2XUc4DADiXHPgs2cK+51ES1caCUxHTrah4h8fTfYG5EB4\njhFxXvJfwOpPt9Ncpr16wzxh95MmV8sY9bPa4Sq5n1XLIN5y5OiJGd6QXwE+j+a8\nOMLLXlathiEK+n6/X3+h/6Xu3NAuTJsn3h7a0HjdKdeR3+ZuZrvWIig3q8XufvDN\nMTtHfxPCxgSFG+owh1wFiq8Q1vlE95xnLrpxD3TBT/GBtPsCdJKRFbr+yTBOttcb\nmUOUzOp6oXC/49J1Y9fVuZjzJfqzZ5x6eRXO1Si/4QKBgQDw1yOqAmsABoh1cRIi\nZfkT3gaS3Kq3X78d8Sugekdugd88kpxotzGiS3DE9E60623Zbt6jxqezwxIBZsOh\nIHzDnKT+JLS6rP3glAThq54DDecOdogvxXWghS+72QkS61YOSRoKIMPX3iYeRAq5\nKwYnRhEMaR1K4uEVi+3RI3MH4QKBgQDJ+lrfglFYYx3ZFhjoKnG95Rc2yB1Rr7AZ\n4UH+JiG54+Rm85W0b2Epp5zs1gqcS2xx3hZGN34z5OqgJduBHRug+5DcSEOytG9R\nZlJ91QETNBZBjM/3bsQz5eksz4j+CgV11e6Dmg3iXOdLd1lPCgt/LTM6BdzAcWpp\nUqy+OsJ6BQKBgBX1F1xNoiG7dr92UpfuQhosmN7U3X+gbBU3wql73H1Xu9mS6E2n\nvg+03xAl0fMur7IuKIA4AVwjQcX874MGKjnPUz+UayHF2dOayyMj+WD/6HvqFJp+\nXy4GVobCz8/4wrzEr2oS+Kf6qfECdRPSt1nnSnCeOLx2GN1VB5aUq80hAoGAOEjb\nAAQZ1Q6x56f/wtrpHWj04iA8A2J5KY0bTc6kgV/fa00f/8s2AVyjH2C6Tjm6e7TO\n8jxOn/l/5KcIF1/cLi1MfgZpTyh3CPEBte0gwpA2T4gFAEfOx0Ofigw/ecOjJ+Y4\n9FV+3wDSt7YHnj4HXCZlaxrtHHe+lqEiYFSRk10CgYA67ZYfmaYw1qdRklEZeI6A\nvlSPlWfx2gp+azYlTreAxHF6FtiMW+wjiB0+HkgAsPnUZ599+KbCWgeHuumq7blx\nAZtPSusOUF91CMwDJCAfdj8lfKgdu6GBebNEe9V23nKiBFaTLAENeWAjB1qtGr8c\nd+9uI1ttV7PPRXLNoW/enQ==\n-----END PRIVATE KEY-----\n",
  "client_email": "bcp-295@bcp-21685.iam.gserviceaccount.com",
  "client_id": "105984466148242757454",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/bcp-295%40bcp-21685.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}


class Pairing:
    def __init__(self, p1_game_result: int, p2_game_result: int, p1_army_id: str, p2_army_id: str):
        self.p1_game_result = p1_game_result
        self.p2_game_result = p2_game_result
        self.p1_army_id = p1_army_id
        self.p2_army_id = p2_army_id
        self.p1_faction = None
        self.p2_faction = None

    def __str__(self):
        return f"{self.p1_army_id} vs {self.p2_army_id}"

    def __repr__(self):
        return self.__str__()


def get_pairings():
    session = requests.Session()
    url = "https://pnnct8s9sk.execute-api.us-east-1.amazonaws.com/prod/pairings?limit=100&eventId=Y3UK4LX1RH&pairingType=Pairing&expand[]=player1&expand[]=player2&expand[]=player1Game&expand[]=player2Game"
    headers = {
        "client-id": "6avfri6v9tgfe6fonujq07eu9c",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/113.0"
    }

    r = requests.get(url, headers=headers)

    yield r.json()["data"]

    next_key = r.json()["nextKey"]

    while True:
        next_page = session.get(url, headers=headers, params={"nextKey": next_key})
        previous_key = next_key
        next_key = next_page.json()["nextKey"]
        if 'offset' in next_key or previous_key == next_key:
            break
        yield next_page.json()["data"]


def get_players():
    session = requests.Session()
    url = "https://pnnct8s9sk.execute-api.us-east-1.amazonaws.com/prod/players?limit=100&eventId=Y3UK4LX1RH&expand%5B%5D=army&expand%5B%5D=subFaction&expand%5B%5D=character&expand%5B%5D=team&expand%5B%5D=user"
    headers = {
        "client-id": "6avfri6v9tgfe6fonujq07eu9c",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/113.0"
    }

    r = requests.get(url, headers=headers)

    yield r.json()["data"]

    next_key = r.json()["nextKey"]

    while True:
        next_page = session.get(url, headers=headers, params={"nextKey": next_key})
        next_key = next_page.json()["nextKey"]
        if len(next_page.json()["data"]) == 0:
            break
        yield next_page.json()["data"]


if __name__ == '__main__':
    pairings: [Pairing] = []

    for pairing_page in get_pairings():
        for pairing in pairing_page:
            try:
                pairings.append(Pairing(
                    p1_game_result=pairing["metaData"]["p1-gameResult"],
                    p2_game_result=pairing["metaData"]["p2-gameResult"],
                    p1_army_id=pairing["player1"]["armyId"],
                    p2_army_id=pairing["player2"]["armyId"],
                ))
            except KeyError:
                pass

    factions = defaultdict(str)

    for player_page in get_players():
        for player in player_page:
            try:
                factions[player["army"]["id"]] = player["army"]["name"]
            except KeyError:
                pass

    for pairing in pairings:
        pairing.p1_faction = factions[pairing.p1_army_id]
        pairing.p2_faction = factions[pairing.p2_army_id]

    results = dict()

    for faction in factions:
        results[factions[faction]] = {}
        for faction2 in factions:
            results[factions[faction]][factions[faction2]] = {'W': 0, 'D': 0, 'L': 0}

    for pairing in pairings:
        if int(pairing.p1_game_result) == 2:
            results[pairing.p1_faction][pairing.p2_faction]['W'] += 1
            results[pairing.p2_faction][pairing.p1_faction]['L'] += 1
        elif int(pairing.p1_game_result) == 0:
            results[pairing.p1_faction][pairing.p2_faction]['L'] += 1
            results[pairing.p2_faction][pairing.p1_faction]['W'] += 1
        else:
            results[pairing.p1_faction][pairing.p2_faction]['D'] += 1
            results[pairing.p2_faction][pairing.p1_faction]['D'] += 1

    wb = Workbook("results.xlsx")
    ws = wb.add_worksheet()

    for i, faction1_name in enumerate(results, start=2):
        ws.write(i, 1, faction1_name)
        ws.write(1, i, faction1_name)
        # sh.sheet1.update_cell(i, 1, faction1_name)
        # sh.sheet1.update_cell(1, i, faction1_name)

    for i, faction1 in enumerate(results, start=2):
        for j, faction2 in enumerate(results[faction1], start=2):
            ws.write(i, j, f"{results[faction1][faction2]['W']}-{results[faction1][faction2]['D']}-{results[faction1][faction2]['L']}")
            # sh.sheet1.update_cell(i, j, results[faction1][faction2])

    wb.close()

    gc = gspread.service_account_from_dict(credentials)
    sh = gc.open("BCP").get_worksheet(0)
    sh.clear()
    df = pd.read_excel("results.xlsx")
    df = df.applymap(lambda x: str(x) if isinstance(x, float) else x)
    sh.insert_rows(df.values.tolist())

