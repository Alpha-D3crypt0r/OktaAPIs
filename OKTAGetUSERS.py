import requests,json
from termcolor import colored
from colorama import Fore
from colorama import Style
import xlrd
loc = ("OktaTest.xls")
wb=xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)
rows=sheet.nrows
print(f"{Fore.GREEN}Kraken's API Get Request Okta Script{Style.RESET_ALL}")
#Reqests library handles the API calls
tenent = "dev-587547.okta"
token = "002x_RGs87sZcSVrsSWI9haeokHyCH1TjqUL1vgQ1L"

def main():
    #REMOVE THE COMMENTS IN THE BELOW 3LINES to not use the Excel Sheet
    #login = input("Enter the username\nOR Enter QUITMOFO to quit ")
    #if login.upper() == "QUITMOFO":
    #    quit()
    i=0
    for i in range(rows):
        content=getUser(sheet.cell_value(i,0))
        print("The UserID for {} is {}: ".format(sheet.cell_value(i,0),content['id']))

#GET USER FUNCTION
def getUser(login):
    my_headers = {"Authorization":"SSWS"+token, "Accept":"application/json", "Content-Type":"application/json"}
    response=requests.get('https://{}.com/api/v1/users?filter=profile.login%20eq%20"{}"'.format(tenent,login) , headers=my_headers)
    #print(r.text) r.text contains the real response fromt the API Server. r only contains the response type eg. 200 404
    #print(r)  #Print response
    if response.status_code == requests.codes['ok']:
        result=json.loads(response.text) #Converts the parsed response in the form of JSON
        try:
            return result[0] #Because the response is a list of dictionaries hence taking the first dict from the list
        except IndexError:
            print(f"{Fore.GREEN}The username is either not correct or the user is not found in the OKTA Database{Style.RESET_ALL}")
            main()
    elif response.status_code == 401:
        print(f'{Fore.GREEN}Check your API Token its not correct{Style.RESET_ALL}')
        main()

main()
