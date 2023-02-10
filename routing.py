from classes import *
import xlwings as xw


# loading required data
static_data = StaticData()
lst_of_bom_obj = []
lst_of_products = []
lst_of_route_df_before = []


# this is how to run python in excel macro
# RunPython "import RouteSheet; RouteSheet.bom_to_route()"

# this will get items in the main sheet in route excel
@xw.func
def main():
    # loading excel handler class
    excelHandler = ExcelHandler()

    # load excel workbook main sheet
    wb = xw.Book.caller()
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
    wb = xw.Book.caller()
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
