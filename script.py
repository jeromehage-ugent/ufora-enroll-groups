import requests, json, time
import numpy as np
import pandas as pd

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# general config
host = 'https://ufora.ugent.be'
orgUnitId = 1025023
groupCategoryId = 170328 # you need to first create the group category manually

# input file
grouplist_path = 'GroupList.xlsx'
grouplist_col = 'Group' # column in spreadsheet with the group name


#### BEGIN

service = Service('chromedriver.exe') # download an updated version of this
options = Options()
options.add_argument('start-maximized')
driver = webdriver.Chrome(service = service, options = options)
action = ActionChains(driver)

## ufora login
print('ufora login')

url = 'https://ufora.ugent.be/d2l/home/{}'.format(orgUnitId)
driver.get(url)

## get session cookies
print('get session cookies')

# used later to read data using API GET requests
# NOTE: UGent does not allow registering new Oauth2 apps
# so I cannot get a bearer token, and API writes don't work (POST/PUT/DELETE)
cookies = {
    'd2lSessionVal': '',
    'd2lSecureSessionVal': '',
    }
for c in cookies:
    v = WebDriverWait(driver, timeout = 120).until(lambda d: d.get_cookie(c))['value']
    cookies[c] = v


### get data from ufora API

## check API version
print('check API version')

path = '/d2l/api/versions/'
r = requests.request('GET', host + path)
data = r.json()

le_version = 1.0
lp_version = 1.0
for d in data:
    if d['ProductCode'] == 'le':
        le_version = d['LatestVersion']
    if d['ProductCode'] == 'lp':
        lp_version = d['LatestVersion']

## get class list from ufora
print('get class list from ufora')

params = {
    'version': le_version,
    'orgUnitId': orgUnitId,
    }
path = '/d2l/api/le/{version}/{orgUnitId}/classlist/'
url = host + path.format(**params)

query_params = {
    'onlyShowShownInGrades': 'true', # doesn't work
    }
if len(query_params) > 0:
    url += '?' + '&'.join(['{}={}'.format(k, v) for k, v in query_params.items()])

r = requests.request(
    'GET', url,
    cookies = cookies,
    )
data = r.json()

classlist = pd.DataFrame(data)

## get list of groups in category
print('get list of groups in category')

params = {
    'version': lp_version,
    'orgUnitId': orgUnitId,
    'groupCategoryId': groupCategoryId,
    }
path = '/d2l/api/lp/{version}/{orgUnitId}/groupcategories/{groupCategoryId}/groups/'
url = host + path.format(**params)

r = requests.request(
    'GET', url,
    cookies = cookies,
    )
data = r.json()

groups = pd.DataFrame(data)

### data processing

classlist['Identifier'] = classlist['Identifier'].astype(np.int64)
classlist['OrgDefinedId'] = classlist['OrgDefinedId'].astype(np.int64)

## read target group list from the file and merge
print('read target group list from the file and merge')

col = grouplist_col

grouplist = pd.read_excel(grouplist_path)
grouplist['grp'] = grouplist.loc[grouplist[col].notna(), col].astype(np.int64)
grouplist.dropna(subset = 'grp', inplace = True)
grouplist['OrgDefinedId'] = grouplist['OrgDefinedId'].astype(np.int64)
grouplist = grouplist[['OrgDefinedId', 'grp']]

classlist = pd.merge(classlist, grouplist, on = 'OrgDefinedId', how = 'inner')

# converts 'Group 3' to 3
def fixname(s):
    return int(s.split(' ')[-1])

groups['grp'] = groups['Name'].apply(fixname)
groups = groups[['grp', 'GroupId']]

classlist = pd.merge(classlist, groups, on = 'grp', how = 'inner')


### web automation

## enrolling students
print('enrolling students')

path = 'https://ufora.ugent.be/d2l/lms/group/group_enroll.d2l?ou={orgUnitId}&categoryId={groupCategoryId}'
params = {
    'orgUnitId': orgUnitId,
    'groupCategoryId': groupCategoryId,
    }
url = path.format(**params)
driver.get(url)

def click_all_checkboxes():
    
    # get the list of all checkboxes
    checkboxes = driver.find_elements(By.CLASS_NAME, 'd2l-checkbox')
    enrol_cbx = {}
    for cbx in checkboxes:
        onclick = cbx.get_attribute('onclick')
        lookup = 'EnrollmentChange('
        if lookup in onclick:
            i = onclick.find(lookup) + len(lookup)
            j = i + onclick[i:].find(')')
            k = onclick[i: j].replace(' ', '')
            enrol_cbx[k] = cbx

    # enroll users in groups by clicking the right checkboxes
    for idx, student in classlist.iterrows():
        groupId = student['GroupId']
        Identifier = student['Identifier']
        key = '{},{}'.format(groupId, Identifier)
        if key in enrol_cbx:
            cbx = enrol_cbx[key]
            #action.move_to_element(cbx).perform()
            #clickable = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(cbx))
            #clickable.click()
            driver.execute_script('arguments[0].click();', cbx)
            classlist.loc[idx, 'clicked'] = True

while True:
    # wait until page is loaded
    checkbox_show = EC.presence_of_element_located((By.CLASS_NAME, 'd2l-checkbox'))
    WebDriverWait(driver, timeout = 120).until(checkbox_show)
    # enroll users
    click_all_checkboxes()
    # scroll to the bottom
    driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
    # click the next button
    buttons = driver.find_elements(By.TAG_NAME, 'd2l-button-icon')
    b_icons = [b.get_attribute('icon') for b in buttons]
    if 'tier1:chevron-right' in b_icons:
        b_next = buttons[b_icons.index('tier1:chevron-right')]
        b_next.click()
    else:
        break

print('ENROLLED', classlist['clicked'].sum(), '/', len(classlist))

missed = classlist[classlist['clicked'] == False]
if len(missed) > 0:
    print('Missing:')
    for idx, student in missed.iterrows():
        print(student['Identifier'], student['DisplayName'])

classlist.to_csv('classlist.csv')
