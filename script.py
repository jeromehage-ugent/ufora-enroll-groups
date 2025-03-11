import requests, json, time
import numpy as np
import pandas as pd

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

def fix_pandas_float_to_int(value):
    # fix pandas sometimes converting integers to floats
    try:
        value = float(value)
        if int(value) == value:
            value = int(value)
    except (ValueError, TypeError):
        pass
    return value


## general config
host = 'https://ufora.ugent.be'
orgUnitId = 1025023

sort_group_names = True # create groups in alphabetical order

# input file
grouplist_path = 'GroupList.xlsx'
student_id_col = 0 # first column

def group_category_formatter(header):
    """Runs over column headers. Replace with your own."""
    # skip these columns
    skip = ['naam', 'name', 'voornaam', 'first name', 'last name', 'uuid']
    if header.lower().strip() in skip:
        return None
    grp_category = 'test_{}'.format(header)
    return grp_category

def group_name_formatter(value):
    """Runs over column values. Replace with your own."""
    if pd.isna(value):
        return value
    value = fix_pandas_float_to_int(value)
    grp_name = 'grp_{}'.format(value)
    return grp_name



## BEGIN

# read grouplist spreadsheet
print('process group list')
grouplist = pd.read_excel(grouplist_path)

# rename student ID column to OrgDefinedId
sid_col = grouplist.columns[student_id_col]
grouplist.rename({sid_col: 'OrgDefinedId'}, axis = 1, inplace = True)
grouplist = grouplist[grouplist['OrgDefinedId'].notna()]
grouplist['OrgDefinedId'] = grouplist['OrgDefinedId'].astype(np.int64)

# convert column headers to group category names
renamer = {c: group_category_formatter(c) for c in grouplist.columns if c != 'OrgDefinedId'}
todrop = [c for c, new_c in renamer.items() if new_c is None]
grouplist.drop(todrop, axis = 1, inplace = True)
renamer = {c: new_c for c, new_c in renamer.items() if new_c is not None}
grouplist.rename(renamer, axis = 1, inplace = True)
gc_names = list(renamer.values()) # group category names

# convert cell values to group names
for gc in gc_names:
    grouplist[gc] = grouplist[gc].apply(group_name_formatter)

# get number of groups in each category
gc_values = {}
gc_count = {}
for gc in gc_names:
    gc_vals = grouplist[gc].dropna().unique().tolist()
    if sort_group_names:
        gc_vals = sorted(gc_vals)
    gc_values[gc] = gc_vals
    gc_count[gc] = len(gc_vals)



## web automation setup

service = Service('chromedriver.exe') # download an updated version of this
options = Options()
options.add_argument('start-maximized')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(service = service, options = options)
action = ActionChains(driver)

# ufora login
print('ufora login')

url = 'https://ufora.ugent.be/d2l/home/{}'.format(orgUnitId)
driver.get(url)

# get session cookies
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



## API setup

# check API version
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

def api_get_classlist():
    
    params = {
        'version': le_version,
        'orgUnitId': orgUnitId,
        }
    path = '/d2l/api/le/{version}/{orgUnitId}/classlist/'
    url = host + path.format(**params)

    r = requests.request(
        'GET', url,
        cookies = cookies,
        )
    data = r.json()

    return pd.DataFrame(data)

def api_get_group_list(groupCategoryId):
    
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

    return pd.DataFrame(data)



## create group categories
print('create group categories and groups')

def waitfor_checkbox_visible():
    # wait until page is loaded
    checkbox_show = EC.presence_of_element_located((By.CLASS_NAME, 'd2l-checkbox'))
    WebDriverWait(driver, timeout = 120).until(checkbox_show)

def set_group_category_name(name):
    """looks for the edit box that has 'name' as its name"""
    editboxes = driver.find_elements(By.CLASS_NAME, 'd2l-edit')
    gcn_box = [e for e in editboxes if e.get_attribute('name') == 'name'][0]
    #ActionChains(driver).move_to_element(inputbox).click(inputbox).send_keys(name).perform()
    gcn_box.clear()
    gcn_box.send_keys(name)
    
def set_group_category_count(count):
    """looks for the input box that has 'gr' in its label (as in group or gropen"""
    inputboxes = driver.find_elements(By.CLASS_NAME, 'd2l-input-number-wc')
    gcc_box = [e for e in inputboxes if 'gr' in e.get_attribute('label')][0]
    gcc_box.send_keys(count)

def click_save():
    """looks for the button with the text 'save' or 'opslaan'"""
    save = ['save', 'opslaan']
    buttons = driver.find_elements(By.CLASS_NAME, 'd2l-button')
    save_button = [e for e in buttons if e.text.lower() in save][0]
    save_button.click()

def click_save_OK():
    """looks for the button with the text 'save' or 'opslaan', clicks pops"""
    save = ['save', 'opslaan']
    buttons = driver.find_elements(By.CLASS_NAME, 'd2l-button')
    save_button = [e for e in buttons if e.text.lower() in save][0]
    save_button.click()
    # wait for any confirmation checkboxes
    ok_button_show = EC.presence_of_element_located((By.CLASS_NAME, 'd2l-dialog-buttons'))
    WebDriverWait(driver, timeout = 10).until(ok_button_show)
    ok_button_clickable = EC.element_to_be_clickable((By.CLASS_NAME, 'd2l-dialog-buttons'))
    ok_button_box = WebDriverWait(driver, timeout = 10).until(ok_button_clickable)
    ok_button = ok_button_box.find_element(By.CLASS_NAME, 'd2l-button')
    ok_button.click()

def set_group_name(name):
    """same behavior as set_group_category_name"""
    set_group_category_name(name)

gc_ids = {}
for gc, count in gc_count.items():
    # go to group category create page
    url = 'https://ufora.ugent.be/d2l/lms/group/category_newedit.d2l?ou={}'.format(orgUnitId)
    driver.get(url)
    # wait until page is loaded
    editbox_clickable = EC.element_to_be_clickable((By.CLASS_NAME, 'd2l-edit'))
    WebDriverWait(driver, timeout = 120).until(editbox_clickable)
    # set name and count
    set_group_category_name(gc)
    set_group_category_count(count)
    # save and wait until group page is loaded
    click_save_OK()
    waitfor_checkbox_visible()
    # extract group category ID from url
    # Note: we could also just make all the groups and read the list of group categories later
    gcid = driver.current_url.split('categoryId=')[1].split('&')[0]
    gc_ids[gc] = gcid
    print('Created group category', gc, gcid)

    # read group ids in group category with API
    group_ids = api_get_group_list(gcid)['GroupId'].values
    for gid, gn in zip(group_ids, gc_values[gc]):
        # go to group edit page
        url = 'https://ufora.ugent.be/d2l/lms/group/group_edit.d2l?ou={}&groupId={}'.format(orgUnitId, gid)
        driver.get(url)
        # wait until page is loaded
        editbox_clickable = EC.element_to_be_clickable((By.CLASS_NAME, 'd2l-edit'))
        WebDriverWait(driver, timeout = 120).until(editbox_clickable)
        # set group name and save
        set_group_name(gn)
        click_save()
        # wait for the page to load
        waitfor_checkbox_visible()
    


## enroll students

def click_all_checkboxes(students):
    """clicks all the appropriate (GroupID, OrgDefinedId) checkboxes"""

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
    for idx, student in students.iterrows():
        key = '{},{}'.format(student['GroupId'], student['Identifier'])
        if key in enrol_cbx:
            cbx = enrol_cbx[key]
            #action.move_to_element(cbx).perform()
            #clickable = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(cbx))
            #clickable.click()
            driver.execute_script('arguments[0].click();', cbx)
            students.loc[idx, 'clicked'] = True

def pages_find_selectors():
    """look for page selector element containing '1 of' or '1 van'"""
    selectors = driver.find_elements(By.CLASS_NAME, 'd2l-select')
    of = ['of', 'van']
    page_selectors = [s for s in selectors if any(['1 {}'.format(o) in s.text for o in of])]
    return page_selectors

def pages_get_count():
    """check how many pages can be selected"""
    selectors = pages_find_selectors()
    if len(selectors) == 0:
        # no page selector means a single page
        return 1
    else:
        select = Select(selectors[0])
        return len(select.options)

def pages_goto(page_nb = 1):
    """click page selector option by index"""
    selectors = pages_find_selectors()
    if len(selectors) != 0:
        select = Select(selectors[0])
        select.select_by_index(page_nb - 1)

def pages_get_current():
    """get index of current selected page"""
    selectors = pages_find_selectors()
    if len(selectors) != 0:
        select = Select(selectors[0])
        return select.options.index(select.first_selected_option) + 1

def perpage_pick_highest():
    """look for page selector element containing '1 of' or '1 van'"""
    selectors = driver.find_elements(By.CLASS_NAME, 'd2l-select')
    perpage_selector = [s for s in selectors if 'per pag' in s.text][0]
    select = Select(perpage_selector)
    nperpage = [int(o.text.split(' per pag')[0]) for o in select.options]
    select.select_by_index(np.argmax(nperpage))




# get class list from ufora
print('get class list from ufora')
classlist = api_get_classlist()
classlist['Identifier'] = classlist['Identifier'].astype(np.int64)
classlist['OrgDefinedId'] = classlist['OrgDefinedId'].astype(np.int64)

# get list of groups in category
print('get list of groups in category')
for gc, gcid in gc_ids.items():
    groups = api_get_group_list(gcid)
    groups = groups[['Name', 'GroupId']]
    # get list of students in this group category
    gc_students = grouplist.loc[grouplist[gc].notna(), ['OrgDefinedId', gc]]
    merged = pd.merge(groups, gc_students, left_on = 'Name', right_on = gc, how = 'outer')
    gc_students = merged[['OrgDefinedId', 'GroupId']]
    # add internal Identifier along with OrgDefinedId
    gc_student_list = pd.merge(classlist, gc_students, on = 'OrgDefinedId', how = 'inner')

    # enrolling students
    print('enrolling students in', gc)

    # go to enrolling students page
    url = 'https://ufora.ugent.be/d2l/lms/group/group_enroll.d2l?ou={}&categoryId={}'.format(orgUnitId, gcid)
    driver.get(url)
    # wait until page is loaded
    waitfor_checkbox_visible()
    # pick most students per page
    perpage_pick_highest()
    waitfor_checkbox_visible()

    pages_visited = []
    while True:
        # select all the correct checkboxes
        click_all_checkboxes(gc_student_list)
        # scroll to the bottom
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')

        # check for multiple pages
        page_cnt = pages_get_count()
        if page_cnt == 1:
            click_save()
            waitfor_checkbox_visible()
            break
        else:
            pages_visited += [pages_get_current()]
            print(pages_visited)
            if len(np.unique(pages_visited)) == page_cnt:
                click_save()
                waitfor_checkbox_visible()
                break
            else:
                # below is kind of a stupid workflow but unfortunately you have to save on every page
                click_save()
                waitfor_checkbox_visible()
                # go back to previous page
                driver.execute_script('window.history.go(-1)')
                waitfor_checkbox_visible()
                # select another page
                not_visited = [i + 1 for i in range(page_cnt) if i + 1 not in pages_visited]
                pages_goto(not_visited[0])
                waitfor_checkbox_visible()

    print('ENROLLED', gc_student_list['clicked'].sum(), '/', len(gc_student_list))

    missed = gc_student_list[gc_student_list['clicked'] == False]
    if len(missed) > 0:
        print('Category', gc, gcid, 'missing:')
        for idx, student in missed.iterrows():
            print(student['Identifier'], student['DisplayName'])

    gc_student_list.to_csv('students_{}.csv'.format(gcid))

input('Done, press enter to leave :-)')
