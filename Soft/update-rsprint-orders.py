import requests
import sys
import os.path

companyNamesToId = { 'go4d': 3, 'superfeet': 4 }

if len(sys.argv) != 3:
    print("Error : Expected the company-name and csv file as arguments")
    sys.exit()

companyString = sys.argv[1]
if companyString not in companyNamesToId.keys():
    print("Error : Could not recognize company " + companyString + "\nSupported names: " + ' '.join(companyNamesToId.keys()))
    sys.exit()
companyId = companyNamesToId[companyString]

csvFile = sys.argv[2]
if os.path.isfile(csvFile) is False: 
    print("Error : Could not find the file " + csvFile)
    sys.exit()

url = 'https://rsscan-identityserver.azurewebsites.net/connect/token'
data = {
    'client_id':'rsprint-jobs', 
    'client_secret': 'gVpXze8L27WG419jMkYqxADNZIdQtFoPKfOhb506lsSw3nEurHCvcUTJRyim',
    'scope':'customer-service:create',
    'grant_type':'client_credentials'
    }
response = requests.post(url, data=data)
responseAsJson = response.json()
accessToken = responseAsJson['access_token']


url = 'https://rsscan-customerservice.azurewebsites.net/api/v1/rsprint-orders'
headers = {'Authorization': 'Bearer ' + accessToken} 
files = {'CsvFile': open(csvFile, 'rb')}
data = {'CompanyId': companyId}

response = requests.post(url, headers=headers, data=data, files=files)
print(response)
print(response.content)
response.raise_for_status()