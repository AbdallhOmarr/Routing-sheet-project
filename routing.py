from classes import *
import xlwings as xw
import warnings
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
    global all_route_df

    # get routing after
    sheet = wb.sheets["Item1"]
    route = sheet.range("A3:BA203").options(
        pd.DataFrame, expand='table',  index=False).value

    all_route_df = pd.concat([all_route_df, route])

    # route.dropna(subset=["dept1"], inplace=True)

    # need to remove the last one if the same button clicked again tbs
    lst_of_route_df_after.append(route)

    for product in lst_of_products[0].copy():
        print(product.code)
        product_route = route[route["item code"] == float(product.code)]
        product_route.dropna(
            subset=["std route", 'copy route', 'dept1'], how='all', inplace=True)

        if pd.notna(product_route["std route"].to_list()[0]):
            print(f"product code:{product.code} has a std route")

            product.std_route = True
            product.get_route(product_route)

        if len(product_route) == 0:
            print(f"product code:{product.code} has no route")
            lst_of_products[0].pop(lst_of_products[0].index(product))
            continue

        product.check_copy_route(product_route)
        if pd.notna(product.copy_route):
            print(f"product code:{product.code} is copy route")
            product_route = all_route_df[all_route_df["item code"]
                                         == product.copy_route]
            product.get_route(product_route)
        else:
            print(f"duplicated:{product.code}")
            product.get_route(product_route)

        product.assign_process()

    routing = Routing(lst_of_products[0])

    sheet = wb.sheets["r1"]
    sheet.range("B4").options(index=False, expand='table',
                              header=False).value = routing.get_wip_data()
    sheet.range("B:B").number_format = '0'

    sheet.range("G4").options(index=False, expand='table',
                              header=False).value = routing.get_operation_data()
    sheet.range("G:G").number_format = '0'

    sheet.range("L4").options(index=False, expand='table',
                              header=False).value = routing.get_resource_data()
    sheet.range("L:L").number_format = '0'
