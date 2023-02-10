from classes import *
import xlwings as xw


# loading required data
static_data = StaticData()
lst_of_bom_obj = []
lst_of_products = []
lst_of_route_df_before = []
lst_of_route_df_after = []

# this is how to run python in excel macro
# RunPython "import RouteSheet; RouteSheet.bom_to_route()"

# this will get items in the main sheet in route excel


# wb caller declared once
wb = xw.Book.caller()


@xw.func
def main():
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
def get_route_data1():
    sheet = wb.sheets["Item1"]
    parent = sheet.range("A2").value
    route = lst_of_route_df_before[0]
    route["item code"] = pd.to_numeric(route["item code"], errors="ignore")
    # route = route.sort_values(by=["item code"], ascending=False)

    # route is loaded and ready

    # route to excel
    sheet.range("A4").options(
        index=False, expand='table', header=False).value = route
    sheet.range("A:A").number_format = '0'
    sheet.range("c:c").number_format = '0'

@xw.func
def get_item_data1():
    # get routing after
    sheet = wb.sheets["Item1"]
    route = sheet.range("A3:BA203").options(
        pd.DataFrame, expand='table',  index=False).value
    route.dropna(subset= ["dept1"],inplace=True)

    # need to remove the last one if the same button clicked again tbs
    lst_of_route_df_after.append(route)
    for product in lst_of_products[0]:
        product_route = route[route["item code"]==float(product.code)] 
        product.get_route(product_route)
        print(f"product code:{product.code}")
        print(f"product route: \n{product.route_processed}")
        product.assign_process()

    routing = Routing(lst_of_products[0])

    sheet = wb.sheets["r1"]
    sheet.range("B4").options(index=False, expand='table', header=False).value = routing.get_wip_data()
    sheet.range("B:B").number_format = '0'

    sheet.range("G4").options(index=False, expand='table', header=False).value = routing.get_operation_data()
    sheet.range("G:G").number_format = '0'

    sheet.range("L4").options(index=False, expand='table', header=False).value = routing.get_resource_data()
    sheet.range("L:L").number_format = '0'

@xw.func
def get_route_data2():
    sheet = wb.sheets["Item2"]
    parent = sheet.range("A2").value
    route = lst_of_route_df_before[1]
    route["item code"] = pd.to_numeric(route["item code"], errors="ignore")
    # route = route.sort_values(by=["item code"], ascending=False)

    # route is loaded and ready

    # route to excel
    sheet.range("A4").options(
        index=False, expand='table', header=False).value = route
    sheet.range("A:A").number_format = '0'
    sheet.range("c:c").number_format = '0'

@xw.func
def get_item_data2():
    # get routing after
    sheet = wb.sheets["Item2"]
    route = sheet.range("A3:BA203").options(
        pd.DataFrame, expand='table',  index=False).value
    route.dropna(subset= ["dept1"],inplace=True)

    # need to remove the last one if the same button clicked again tbs
    lst_of_route_df_after.append(route)
    for product in lst_of_products[1]:
        product_route = route[route["item code"]==float(product.code)] 
        product.get_route(product_route)
        print(f"product code:{product.code}")
        print(f"product route: \n{product.route_processed}")
        product.assign_process()

    routing = Routing(lst_of_products[1])

    sheet = wb.sheets["r2"]
    sheet.range("B4").options(index=False, expand='table', header=False).value = routing.get_wip_data()
    sheet.range("B:B").number_format = '0'

    sheet.range("G4").options(index=False, expand='table', header=False).value = routing.get_operation_data()
    sheet.range("G:G").number_format = '0'

    sheet.range("L4").options(index=False, expand='table', header=False).value = routing.get_resource_data()
    sheet.range("L:L").number_format = '0'
