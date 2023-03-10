from classes import *
import xlwings as xw
import warnings
from essentials import *
warnings.filterwarnings("ignore")


# loading required data
static_data = StaticData()
lst_of_bom_obj = []
lst_of_products = []
lst_of_route_df_before = []
lst_of_route_df_after = []
all_route_df = pd.DataFrame()
# this is how to run python in excel macro
# RunPython "import RouteSheet; RouteSheet.bom_to_route()"

# this will get items in the main sheet in route excel


# wb caller declared once
wb = xw.Book.caller()


@xw.func
def main():
    global lst_of_bom_obj
    global lst_of_products
    global lst_of_route_df_before
    global lst_of_route_df_after
    global all_route_df
    global wb
    wb = xw.Book.caller()
    lst_of_bom_obj = []
    lst_of_products = []
    lst_of_route_df_before = []
    lst_of_route_df_after = []
    all_route_df = pd.DataFrame()
    # loading excel handler class
    excelHandler = ExcelHandler()

    # load excel workbook main sheet
    items = wb.sheets["main"].range("A1:c11").options(
        pd.DataFrame, expand='table', index=False).value
    items = items.dropna()
    # get parents first
    excelHandler.get_parent_items(items)

    # this will get bom data after loading parents
    boms_df = excelHandler.get_bom_data()

    # create loop for each bom to add it to a new sheet
    for bom_df in boms_df:
        bom_obj = Bom(bom_df)
        products = bom_obj.get_lst_of_products()
        lst_of_bom_obj.append(bom_obj)
        lst_of_products.append(products)

    # get lst of route df before filling
    for bom in lst_of_bom_obj:
        lst_of_route_df_before.append(bom.get_route_df())

    # then add each df in route df into a new separte file


@xw.func
def get_route_data():
    sheet = wb.sheets.active
    sheet.range("A4:BC205").value = ""
    sheet.range("A1:G205").color = None
    for col in sheet.range("H1:be205").columns:
        if (col.column - 3) % 5 == 0:
            col.color = (0, 0, 0)  # Set skipped columns to black
        else:
            col.color = None  # Set non-skipped columns to default color

    parent = sheet.range("A2").value
    for bom in lst_of_bom_obj:
        if float(bom.top_parent) == float(parent):
            route = bom.get_route_df()

    route["item code"] = pd.to_numeric(route["item code"], errors="ignore")
    # route = route.sort_values(by=["item code"], ascending=False)

    # route is loaded and ready

    # route to excel
    sheet.range("A4").options(
        index=False, expand='table', header=False).value = route
    sheet.range("A:A").number_format = '0'
    sheet.range("c:c").number_format = '0'


@xw.func
def get_item_data():
    global all_route_df
    active_sheet = wb.sheets.active
    active_sheet_name = active_sheet.name
    active_sheet.range("B4:R10000").value = ""
    active_sheet.range("B4:R10000").color = None

    # get routing after
    sheet = wb.sheets[f"Item{active_sheet_name[-1]}"]
    route = sheet.range("A3:BD203").options(
        pd.DataFrame, expand='table',  index=False).value

    parent = sheet.range("A2").value
    if parent:
        for bom in lst_of_bom_obj:
            if float(bom.top_parent) == float(parent):
                products = bom.get_lst_of_products()

    all_route_df = pd.concat([all_route_df, route])

    # route.dropna(subset=["dept1"], inplace=True)

    # need to remove the last one if the same button clicked again tbs
    lst_of_route_df_after.append(route)

    for product in products.copy():
        print(f"-------------{product.code}------------")
        try:
            product_route = route[route["item code"] == float(product.code)]
        except:
            product_route = route[route["item code"] == product.code]

        product_route.dropna(
            subset=["std route", 'copy route', 'dept1'], how='all', inplace=True)
        if len(product_route) > 0:
            if pd.notna(product_route["std route"].to_list()[0]):

                product.std_route = True

        if len(product_route) == 0:
            products.pop(products.index(product))
            continue

        product.check_copy_route(product_route)
        if pd.notna(product.copy_route):
            product_route = all_route_df[all_route_df["item code"]
                                         == product.copy_route]
            product.get_route(product_route)
        else:
            product.get_route(product_route)

        product.assign_process()

    routing = Routing(products)

    sheet = wb.sheets.active
    sheet.range("B4").options(index=False, expand='table',
                              header=False).value = routing.get_wip_data()
    sheet.range("B:B").number_format = '0'

    sheet.range("G4").options(index=False, expand='table',
                              header=False).value = routing.get_operation_data()
    sheet.range("G:G").number_format = '0'

    sheet.range("L4").options(index=False, expand='table',
                              header=False).value = routing.get_resource_data()
    sheet.range("L:L").number_format = '0'


@xw.func
def to_dataloader():
    all_dl()


@xw.func
def to_mail():
    maill()


@xw.func
def clear_sheets():
    wb = xw.Book.caller()
    sheets = ['Item1', 'Item2', 'Item3', 'Item4', 'Item5']

    for sheet in sheets:
        sheet = wb.sheets[sheet]
        sheet.range("A4:BC205").value = ""
        sheet.range("A1:G205").color = None
        for col in sheet.range("H1:be205").columns:
            if (col.column - 3) % 5 == 0:
                col.color = (0, 0, 0)  # Set skipped columns to black
            else:
                col.color = None  # Set non-skipped columns to default color

    sheets = ['r1', 'r2', 'r3', 'r4', 'r5', "route"]
    for sheet in sheets:
        sheet = wb.sheets[sheet]
        sheet.range("B4:R10000").value = ""
        sheet.range("B4:R10000").color = None

    print("Sheets cleared")


@xw.func
def append():
    append_routing()
