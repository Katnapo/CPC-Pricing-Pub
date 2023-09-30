import json
import requests
import time
from constants import Constants

def constructHTTPComponent(type, pairs):
    returnDict = {}
    for pair in pairs:
        returnDict[pair[0]] = pair[1]

    if type == "B":
        return json.dumps(returnDict)

    else:
        return returnDict

def bcGetRequest(url, attempt=1, debug=False):

    headers = constructHTTPComponent("H", [["Authorization", Constants.BC_KEYS], ["Content-Type", "application/json"]])
    response = requests.request("GET", url, headers=headers)
    if response.status_code != 200:
        if attempt == 3:
            print("Error at bcGetRequest: " + str(response.status_code) + " " + str(response.text))
            print("URL: " + url)
            return None
        time.sleep(0.5)
        return bcGetRequest(url, attempt=attempt+1)

    else:
        realJson = json.dumps(response.json(), indent=4, sort_keys=True)
        if debug == True:
            print(realJson)
        realDict = json.loads(realJson)  # Loads json into a dictionary

        return realDict