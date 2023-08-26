import pandas as pd
import math
#1 Connecting with Excel File
df = pd.read_excel(r'offer_list.xlsx')

#2 Connecting with Shoper
import requests
login = 'login'
haslo = 'password?'

shop = 'https://shopname.com'

s = requests.Session()
response = s.post(shop + '/webapi/rest/auth', auth=(login, haslo))
result = response.json()
token = result['access_token']
s.headers.update({'Authorization': 'Bearer %s' % token})

# Checking if it's working

response = s.get(shop + '/webapi/rest/products/'+'10')
print(response.json())

# Code that downloads the data from Shoper and updates the Excel file with that

#1 function that takes the product id and performs the operation
def updateExcel(offer_id):
    id = int(offer_id)
    
    response = s.get('https://mkowalska.art/webapi/rest/products/'+str(offer_id))
    offer = response.json()
    
    print(int(offer['product_id']) == id)
    
    pl_name = offer['translations']['pl_PL']['name']
    pl_desc = offer['translations']['pl_PL']['description']
    en_name = ''
    en_desc = ''
    
    try:
        en_name = offer['translations']['en_US']['name']
        en_desc = offer['translations']['en_US']['description']
    except:
        pass
    
    print('old offer data from file:')
        
    print(df.loc[id,'title_PL'])
    print(df.loc[id,'title_EN'])
    print(df.loc[id,'desc_PL'])
    print(df.loc[id,'desc_EN'])
    
    df.loc[id,'title_PL'] = pl_name
    df.loc[id,'title_EN'] = en_name
    df.loc[id,'desc_PL'] = pl_desc
    df.loc[id,'desc_EN'] = en_desc
    
    print('new offer data in file:')
    print(df.loc[id,'title_PL'])
    print(df.loc[id,'title_EN'])
    print(df.loc[id,'desc_PL'])
    print(df.loc[id,'desc_EN'])


# Checking if above function works

updateExcel(30)

# loop - to get all translations

for i in df.index:
    print(i)
    try:
        updateExcel(i)
    except:
        continue

#save file
df.to_excel('offer_list.xlsx')

# loop - to get translations by ids [97-99]

for i in range(97,99):
    print(i)
    try:
        updateExcel(i)
    except:
        continue
      

# loop - to get all MISSING translations

for i in df.index:
    print(i)
    try:
        if math.isnan(df.loc[i,'desc_PL']):
            updateExcel(i)
    except:
        continue
