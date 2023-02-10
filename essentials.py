from datetime import datetime
import shutil
import mysql.connector as mysql
import webbrowser as browser
import warnings
import autogui as fa
import math
import pyperclip as pc
import keyboard as kb
import time
from matplotlib.pyplot import table
import xlwings as xw
import pandas as pd
import os
import numpy as np
import pickle
from sklearn.preprocessing import normalize
import random
import pyautogui as pa
import win32com.client as win32


pa.FAILSAFE = True
warnings.filterwarnings("ignore")


def get_files(ext, path):
    print('Getting files with extension: ' + ext)
    path = __location__ + os.sep + path
    print('Looking in: ' + path)
    files = []
    for r, d, f in os.walk(path):
        for file in f:
            if ext in file and "$" not in file:
                files.append(os.path.join(r, file))
                print("file: " + file)

    print('Found ' + str(len(files)) + ' files')
    return files

    # return [os.path.join(path,f) for f in os.listdir(path) if f.endswith(ext)]


def to_excels(df, xls_name):
    '''This method used to debug code'''
    df.to_excel(__location__+f"//xlss//{xls_name}.xlsx")


def get_img(folder_pth):
    '''for downloading bom'''
    lst_of_png = os.listdir(folder_pth)
    lst_of_pngs = [x for x in lst_of_png if x.endswith('.png')]

    n = 0
    while True:
        n += 1
        for img_pth in lst_of_pngs:
            start = pa.locateCenterOnScreen(folder_pth+os.sep+img_pth)

        if n > 10:
            break

        if start is not None:
            return start


def open_url():

    url = R'http://10.10.1.27:8080/ords/acrow_dev/r/acrow-misr-info-center'
    browser.open_new(url)


def login():
    ret = get_img(__location__+os.sep+"\imgs\password")
    if ret is None:
        pa.moveTo(x=909, y=628)
    else:
        pa.moveTo(ret.x, ret.y)
    pa.doubleClick()
    kb.write('Aa#464')
    kb.press_and_release('enter')


def open_lst():
    ret = get_img(__location__+os.sep+R"\imgs\3 lines")
    if ret is None:
        pa.moveTo(25, 126)
    else:
        pa.moveTo(ret.x, ret.y)
    pa.click()


def open_indented_lst():
    ret = get_img(__location__+os.sep+R"\imgs\tech office")
    if ret is None:
        pa.moveTo(x=225, y=396)
    else:
        pa.moveTo(ret.x, ret.y)

    pa.moveRel(150, 0)
    pa.click()
    pa.moveRel(0, 50)
    pa.click()
    ret = get_img(__location__+os.sep+R"\imgs\indented bom")
    if ret is None:
        pa.moveTo(x=99, y=472)
    else:
        pa.moveTo(ret.x, ret.y)

    pa.click()


def add_code(code):
    ret = get_img(__location__+os.sep+R"\imgs\barcode")
    if ret is None:
        pa.moveTo(x=492, y=184)
    else:
        pa.moveTo(ret.x, ret.y)

    pa.moveRel(200, 0)
    pa.click()
    kb.press_and_release('ctrl+a')
    time.sleep(0.2)
    kb.press_and_release('delete')
    code = str(code).split('.')[0]
    kb.write(code)
    pa.press('tab')
    pa.press('tab')
    kb.press_and_release('space')
    pa.doubleClick()


def download_boms():
    time.sleep(3)
    ret = get_img(__location__+os.sep+R"\imgs\action")
    if ret is None:
        pa.moveTo(x=614, y=228)
    else:
        pa.moveTo(ret.x, ret.y)
    pa.click()
    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('down')
    time.sleep(0.2)

    kb.press_and_release('enter')
    time.sleep(0.2)

    kb.press_and_release('right')
    time.sleep(0.2)

    kb.press_and_release('right')
    time.sleep(0.2)

    kb.press_and_release('right')
    time.sleep(0.2)

    kb.press_and_release('tab')
    time.sleep(0.2)

    kb.press_and_release("space")
    time.sleep(0.2)

    # press tab
    kb.press_and_release('tab')
    time.sleep(0.2)

    kb.press_and_release('tab')
    time.sleep(0.2)
    kb.press_and_release('tab')
    time.sleep(0.2)

    kb.press_and_release('enter')


### used in calculating routing rates ###
def calc_laser(l, w, t, n):
    '''This method to calculate laser based on laser table from Purchasing'''
    prm = (2*(l+w))+(15*math.pi*n)
    laser_table = {1: 9, 2:   6.25, 3: 4.25, 4: 3.5, 5: 3.15, 6: 2.95, 8: 2.5, 10: 1.95, 12: 1.5, 14: 1.05, 16: 0.9, 18: 0.75, 20: 0.65, 22: 0.6
                   }

    # convert prm from mm to m
    prm = prm/1000
    # t = str(t)
    # t = t.split('.')[0]

    while True:
        print("laser loop")
        print(f"thickness of the sheet: {t}")
        if t in laser_table.keys():
            break
        print(t)
        t = float(t)
        if t < 22:
            # t = str(t+1)
            t += 1
        if t > 22:
            t = 22
        else:
            # t = str(t-1)
            t += 1

    #     if '.' in t:
    #         t = t.split('.')[0]

    # if '.' in t:
    #     t = t.split('.')[0]
    print(t)

    laser_power = laser_table[t]
    laser_speed = prm/laser_power
    productivity = 60/laser_speed
    return productivity


def ceil(x, s):
    return s * math.ceil(float(x)/s)

# getting boms files
# RunPython "import NewCap; NewCap.all_dl()"


__location__ = os.path.realpath(os.path.join(
    os.getcwd(), os.path.dirname(__file__)))


@xw.func
def bom_to_route():
    print("Running RouteSheet")
    wb = xw.Book.caller()
    print("workbook: ", wb)

    print("Getting RouteSheet")
    items = wb.sheets["Items"].range("A1:c11").options(
        pd.DataFrame, expand='table', index=False).value
    print(items)
    items = items.dropna()
    # open_url()
    # login()
    # open_lst()
    # open_indented_lst()
    # print(items['Items Code'].to_list())
    # for code in items['Items Code']:
    #     print("Downloading code " + str(code))
    #     add_code(code)
    #     time.sleep(1)
    #     download_bom()
    #     print("done downloading code " + str(code))

    print("Getting BOMS")
    boms = get_files(".xlsx", "Boms")
    print("done getting boms and here its: ", boms)

    for i, v in items.iterrows():
        parent_code = v['Items Code']
        parent_desc = v['Item Desc']
        sheet_no = v['no']

        print("items: ", items)
        print("loop on boms")
        print("len(boms): ", len(boms))
        for bom in boms:
            print("bom: ", bom)
            bom = pd.read_excel(bom)
            # if top parent is not in bom, then continue
            print(bom.columns.to_list())
            if "Top Parent" not in bom.columns.to_list():
                continue

            print(parent_code)
            try:
                float(parent_code)
                bom['Top Parent'] = pd.to_numeric(bom['Top Parent'])
            except:
                pass

            print("bom: ", bom)

            if bom.shape[0] == 0:
                continue

            parent_code_bom = bom['Top Parent'].iloc[0]
            if parent_code_bom == parent_code:
                print("-"*50)
                # reading bom

                # filtering bom
                print("filtering bom")
                bom = bom[bom['Comp Item Type'] != 'RESIDUAL']

                bom['Assembly Item'] = pd.to_numeric(
                    bom['Assembly Item'], errors='ignore')

                bom['Component Item'] = pd.to_numeric(
                    bom['Component Item'], errors='ignore')

                print(bom['Assembly Item'])
                sub_assemb_544 = bom[bom['Assembly Item'].astype(
                    str).str.startswith('544')]
                sub_assemb_522 = bom[bom['Assembly Item'].astype(
                    str).str.startswith('522')]
                sub_assembly = sub_assemb_544.append(sub_assemb_522)
                sub_assembly = sub_assembly['Assembly Item'].to_list()
                sub_assembly.append(parent_code)
                # convert all to float

                # sub_assembly = [float(i) for i in sub_assembly]
                assemblies = []
                for sub_item in sub_assembly:
                    try:
                        assemblies.append(float(sub_item))
                    except:
                        assemblies.append(sub_item)
                    print("sub_item: ", sub_item)
                print("sub_assembly: ", assemblies)
                print(bom.info())
                bom['Assembly Item'] = pd.to_numeric(
                    bom['Assembly Item'], errors='coerce')

                route = bom[bom['Assembly Item'].isin(assemblies)]
                print(bom['Assembly Item'])
                route = route[['Component Item', 'Comp Desc', 'Extended Qty', 'Comp Unit Length', 'Comp Unit Width', 'Comp Unit Height', 'Calc Unit Weight',
                               'Related Item', 'Related Desc', 'Related Unit Length', 'Related Unit Width', 'Related Unit Height', 'Related Unit Weight',
                               'Comp Item Status', "Comp Major Category",	"Comp Sub Category",	"Comp Minor Category",	"Comp Item Class",	"Parent Item Class"]]

                route.rename(columns={"Extended Qty": "Qty", "Comp Unit Length": "Length", "Comp Unit Width": "Width", "Comp Unit Height": "Height",
                                      "Calc Unit Weight": "Weight", "Comp Major Category": "Major Category", "Comp Minor Category": "Minor Category",
                                      "Comp Sub Category": "Sub Category", "Comp Item Class": "Item Class", "Comp Item Status": "Status"}, inplace=True)

                sheet_no = str(sheet_no).split(".")[0]
                sheet = wb.sheets[f"Item{sheet_no}"]
                print(route['Parent Item Class'])
                parent = pd.DataFrame(
                    {"Component Item": parent_code, "Comp Desc": parent_desc, "Item Class": route.at[0, 'Parent Item Class']}, index=[0])

                # make route dataframe index starts from 1 instead of 0
                route.reset_index(drop=True, inplace=True)
                route.index += 1
                route = parent.append(route)
                # parent = parent.reindex(columns=route.columns)
                # route = pd.concat([parent, route])
                route['Component Item'] = pd.to_numeric(
                    route['Component Item'], errors='ignore')
                # sort values by index
                # route = route.sort_values(by=route.index)

                # route.sort_values(by=['Component Item'],
                #                   ascending=False, inplace=True)
                route.drop(columns=['Parent Item Class'], inplace=True)
                sheet.range("A1").options(
                    index=False, expand='table').value = route

            else:
                print("no bom for this item")
    # move boms to old boms folder after done processing
    print("moving boms to old boms folder")
    for bom in boms:
        bom1 = bom[:-5] + f"{random.randint(0,99999)}.xlsx"
        bom1 = bom1[:25]+'old boms'+bom1[29:]
        print(bom1)
        print(bom)
        # D:\repo\Newest-Route-App\Boms\indented_bom (1)95831.xlsx
        # D:\repo\Newest-Route-App\Boms\indented_bom (1).xlsx

        shutil.move(bom, bom1)

    print("done moving boms to old boms folder")


# clear sheet

@ xw.func
def clear_sheets():
    wb = xw.Book.caller()
    sheets = ['Item1', 'Item2', 'Item3', 'Item4', 'Item5',
              'Item6', 'Item7', 'Item8', 'Item9', 'Item10']
    for sheet in sheets:
        sheet = wb.sheets[sheet]
        sheet.range("A2:AT150").value = ""
    print("Sheets cleared")

    sheets = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'Route']
    for sheet in sheets:
        sheet = wb.sheets[sheet]
        sheet.range("B3:S10000").value = ""
    print("Sheets cleared")


@ xw.func
def maill():
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    # fil = wb.name
    __location__ = os.path.realpath(os.path.join(
        os.getcwd(), os.path.dirname(__file__)))

    today_date = datetime.today().strftime('%d-%m-%Y')
    df = sheet.range("B1:C10").options(pd.DataFrame, index=False).value
    df.dropna(subset=['Items Code'], inplace=True)

    num = int(sheet.range("H8").value)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "mohamed.gaafer@acrow.co;Mahmoud.Shaban@Acrow.co"
    mail.cc = "Abdallah.Ashry@Acrow.co;fatma.ahmed@Acrow.co;Hossam.Ahmed@Acrow.co;mostafa.dahy@Acrow.co;abdelrahman.amer@Acrow.co;ahmed.abdalkareem@acrow.co;Ahmed.Ashraf@Acrow.co;ahmed.shabana@Acrow.co;ashraf.abdelsatar@Acrow.co;hussien.elsakary@Acrow.co;ibrahim.elsayed@Acrow.co;mahmoud.abdelrasoul@Acrow.co;mahmoud.naser@Acrow.co;Mohamed.Elhussainy@acrow.co;mostafa.dahy@Acrow.co;mostafa.sabry@acrow.co;ramadan.selim@Acrow.co;shady.bakr@Acrow.co;somaya.abdelkader@Acrow.co;Yahia.Hamdy@Acrow.co"

    mail.Subject = f'Finished Routing {num} ({today_date})'
    # mail.Body = 'How are you'
    mail.HTMLBody = "<br>" + df.to_html(index=False, justify='center', formatters={
        'Items Code': lambda x: f"{x:.0f}"})  # this field is optional

    # file_name= __location__ + "\\"+ fil
    # mail.Attachments.Add(file_name)
    mail.Display(True)
    # mail.Send()
    print("Opened")
