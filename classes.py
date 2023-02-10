
# This will contain all requrired product


# imports
import glob
import os
import pandas as pd
import xlwings as xw

class Product:
    def __init__(self, code, description, main_category, sub_category, minor_category, item_type, weight, length, width, thickness, comp_qty, status, raw_material=None, locator=None) -> None:

        # main attirbutes
        self.code = code
        self.description = description
        self.locator = locator
        self.main_category = main_category
        self.sub_category = sub_category
        self.minor_category = minor_category
        self.item_type = item_type

        # physical attributes
        self.weight = weight
        self.length = length
        self.width = width
        self.thickness = thickness

        # manufacturing related attributes
        self.comp_qty = comp_qty

        # status attributes
        self.status = status

        # raw material
        self.raw_material = raw_material

        # product will have processes
        self.lst_of_processes = []

    def get_product_vector(self,):
        # this will return product vector to calculate process cycle time
        product_vector = [self.main_category, self.sub_category, self.minor_category,
                          self.length, self.width, self.thickness, self.weight, self.comp_qty]
        return product_vector

    def assign_process(self, process_matrix):
        # this will add a process obj to the lst of processes
        process = Process(
            process_matrix["process"]["code"], process_matrix["process"]["sequence"])

        # this step should be done once in the main app
        process.get_process_factors("path_to_factors_Sheet")
        process.assign_department("dept_code")
        process.assign_machine("machine_code", "no_of_resource", "res seq")
        process.assign_labor("labor code", "no of resource", "res seq")
        process.calc_rate()
        # after initializing process append it to the lst of processes for the current code
        self.lst_of_processes.append(process)

    def get_locator(self):
        # by getting the last op dept and then getting its wip
        last_process = self.lst_of_processes[-1]
        wip = last_process.wip
        return wip

    def get_wip_data(self):
        # this function will return a dict containing first table data [{wip data}]
        wip_data = {
            "Parts Code": self.code,
            "Description": self.description,
            "WIP": self.get_locator(),
            "Locator": self.get_locator()+".Ground..",
        }
        return wip_data

    def get_operation_data(self):
        # this function will return a list of dict containing second table [{operation data}]

        lst_of_operation_data = []

        for process in self.lst_of_processes:
            operation_data = {
                "Part Code": self.code,
                "Operation Sequence": process.op_seq,
                "Operation Code": process.code,
                "batch_size": process.min_order_qty
            }

            lst_of_operation_data.append(operation_data)

        return lst_of_operation_data

    def get_resource_data(self):
        # this function will return a list of dict containing third table [{resources data}]

        lst_of_resource_data = []

        for process in self.lst_of_processes:
            resource_data = {
                "Part Code": self.code,
                "Operation Sequence": process.op_seq,
                "Resource Sequence": process.machine.res_seq,
                "Resource Code": process.machine.code,
                "Inverse": process.machine.rate,
                "Assigned Units": process.machine.no_of_resource
            }

            lst_of_resource_data.append(resource_data)

            resource_data = {
                "Part Code": self.code,
                "Operation Sequence": process.op_seq,
                "Resource Sequence": process.labor.res_seq,
                "Resource Code": process.labor.code,
                "Inverse": process.labor.rate,
                "Assigned Units": process.labor.no_of_resource
            }
            lst_of_resource_data.append(resource_data)

        return lst_of_resource_data


class Bom:
    def __init__(self, bom_df) -> None:
        # in this class i will get bom excel sheet and extract products and raw material from it.
        # bom_file is the file path for bom
        self.bom_df = bom_df

    def get_lst_of_products(self):
        # - each line will contain a data of a product

        # loop on the dataframe
        # inialize products
        # append to lst_of_products
        # return list for products
        parts_code_start = ["422", "322", "522"]
        lst_of_products = []
        for i, v in self.bom_df.iterrows():
            if (v["Comp Item Type"] == "Part") or (v["Component Item"][:3] in parts_code_start):
                product = Product(v["Component Item"], v["Comp Desc"], v["Comp Major Category"], v["Comp Sub Category"], v["Comp Minor Category"], v["Comp Item Type"],
                                  v["Calc Unit Weight"], v["Comp Unit Length"], v["Comp Unit Width"], v["Comp Unit Height"], v["Comp Qty"], v["Comp Item Status"], v["Related Item"])
                lst_of_products.append(product)
        return lst_of_products

    def get_route_df(self):
        # in this function i will return data to enable user to assign factory, process, machine, no of labors
        lst_of_route_items = []

        parent_code = self.bom_df["Top Parent"].to_list()[0]
        parent_desc = self.bom_df["Parent Description"].to_list()[0]
        lst_of_route_items.append({
            "item code": parent_code,
            "item description": parent_desc,
            "material description": "-"
        })

        # sort bom df by code
        # self.bom_df = self.bom_df.sort_values(by=parent_code)

        for i, v in self.bom_df.iterrows():
            if (v["Component Item"].startswith("RE")):
                continue
            item_dict = {
                "item code": v["Component Item"],
                "item description": v["Comp Desc"],
                "material description": v["Related Desc"]
            }
            lst_of_route_items.append(item_dict)
        route_df = pd.DataFrame(lst_of_route_items)

        return route_df


class Department:
    def __init__(self, code) -> None:
        self.code = code

    def get_wip(self):
        # get sheet for each dept wip
        # Filter on for current dept
        # return wip value
        return "WIP"


class Machine:
    def __init__(self, code, no_of_resource, res_seq) -> None:
        self.code = code
        self.no_of_resource = no_of_resource
        self.res_seq = res_seq

    def assign_rate(self, rate):
        self.rate = rate/self.no_of_resource


class Labor:
    def __init__(self, code, no_of_resource, res_seq) -> None:
        self.code = code
        self.no_of_resource = no_of_resource
        self.res_seq = res_seq

    def assign_rate(self, rate):
        self.rate = rate/self.no_of_resource


class Process:
    def __init__(self, code, op_seq, no_of_cuts, min_order_qty=None, wip=None) -> None:
        self.code = code
        self.op_seq = op_seq
        self.min_order_qty = min_order_qty
        self.wip = wip
        self.no_of_cuts = no_of_cuts

    def get_process_factors(self, factors_sheet):
        # get factors to calc process cycle time
        # factors sheet is a path for excel

        factors_df = pd.read_excel(factors_sheet)
        return factors_df

    def assign_department(self, dept_code):
        self.department = Department(dept_code)

        # after getting dept assign wip
        # value for each dept may be from Department class
        self.wip = self.department.get_wip()

    def assign_machine(self, machine_code, no_of_resource, res_seq):
        self.machine = Machine(machine_code, no_of_resource, res_seq)

    def assign_labor(self, labor_code, no_of_resource, res_seq):
        self.labor = Labor(labor_code, no_of_resource, res_seq)

    def calc_rate(self):
        # in this function i will use product vector and factors_df to calc rate for this process
        # after calc rate
        # assign min order qty
        # value of multiple of 50 near the calc_rate
        self.rate = "rate" / self.no_of_cuts
        self.min_order_qty = "min order qty"
        # assign rate for machine and labor
        self.machine.assign_rate(self.rate)
        self.labor.assign_rate(self.rate)


class Routing:
    def __init__(self) -> None:
        # this class to interact with Excel and Bom and product
        pass

    def get_route_df_before(self):
        # get route df from Bom before filling it from User
        pass

    def get_route_df_after(self):
        # get route df from excel after user assigned everything
        pass

    def get_process_matrix(self):
        # after getting route df process it to provide process matrix for products which will enable product to assign dept, process, machines, labors
        pass

    def get_wip_data(self):
        # this will get wip data for each product
        # aggregate data into a list or dataframe
        pass

    def get_operation_data(self):
        # this will get operation data for each product
        # aggregate data into a list or dataframe
        pass

    def get_resource_data(self):
        # this will get resource data for each product
        # aggregate data into a list or dataframe
        pass


class StaticData:
    def __init__(self):
        self.wb = xw.Book.Caller()
        
    def load_department_excel(self):
        df = self.wb.sheets["department"].range("A1:c100").options(
        pd.DataFrame, expand='table', index=False).value
        df.dropna(inplace=True)
        return df

    def load_process_excel(self, path):
        df = self.wb.sheets["operations"].range("A1:E300").options(
        pd.DataFrame, expand='table', index=False).value
        df.dropna(inplace=True)
        return df

    def load_machines_excel(self, path):
        df = self.wb.sheets["machines"].range("A1:D300").options(
        pd.DataFrame, expand='table', index=False).value
        df.dropna(inplace=True)
        return df

    def load_labors_excel(self, path):
        df = self.wb.sheets["labors"].range("A1:D300").options(
        pd.DataFrame, expand='table', index=False).value
        df.dropna(inplace=True)
        return df

    def load_process_factors_excel(self, path):
        df = self.wb.sheets["rates"].range("A1:D300").options(
        pd.DataFrame, expand='table', index=False).value
        df.dropna(inplace=True)
        return df
    def get_from_dept(self, dept_code):
        filtered_data = self.dept_df[self.dept_df["code"] == dept_code]
        return filtered_data

    def get_from_process(self, process_code):
        filtered_data = self.process_df[self.process_df["code"]
                                        == process_code]
        return filtered_data

    def get_from_machine(self, machine_code):
        filtered_data = self.machine_df[self.machine_df["code"]
                                        == machine_code]
        return filtered_data

    def get_from_labor(self, labor_code):
        filtered_data = self.labor_df[self.labor_df["code"] == labor_code]
        return filtered_data

    def get_from_process_factors(self, id):
        filtered_data = self.process_factors_df[self.process_factors_df["code"] == id]
        return filtered_data


class Dataloader:
    def __init__(self) -> None:
        # this is a class to get tables from excel and load it in oracle data loader
        pass

    def get_tables(self):
        # get tables from excel
        pass

    def load_dataloader(self):
        # laod dl
        pass


class ExcelHandler:
    # this class will be the interface to communicate with Excel
    # will get all the required data and provide all the required data

    def __init__(self) -> None:
        self.cwd = __location__ = os.path.realpath(
            os.path.join(os.getcwd(), os.path.dirname(__file__)))

    def get_parent_items(self, main_df):
        # load items from xlsx
        main_df = main_df.dropna()
        self.parent_items = []
        for i, v in main_df.iterrows():
            self.parent_items.append(v["Items Code"])

        print(self.parent_items)

    def get_last_10_modified_xlsx_files(self, folder_path):
        all_xlsx_files = glob.glob(os.path.join(folder_path, '*.xlsx'))
        all_xlsx_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        return all_xlsx_files[:10]

    def get_bom_data(self):
        # load bom excels and save it in a lst of dataframes
        # return lst of dataframes
        boms_pth = self.cwd + "\\" + "boms"

        # load only xlsx files from this folder
        last_10_modified_xlsx_files = self.get_last_10_modified_xlsx_files(
            boms_pth)

        # filter the last 10 by items in main sheet
        # lst of bom data
        bom_data = []

        for bom in last_10_modified_xlsx_files:
            bom_df = pd.read_excel(bom)
            print(bom_df)
            try:
                if bom_df["Top Parent"].to_list()[0] in self.parent_items:
                    bom_data.append(bom_df)
            except:
                continue
        return bom_data

    def get_route_data_before(self):
        # this will get route data before filling from route class and store it in a list of dataframes
        pass

    def get_route_data_after(self):
        # this will get route data after filling from excel sheet
        pass

    def get_wip_table(self):
        # this will get wip tables and store it in a list of dataframes
        pass

    def get_operation_table(self):
        # this will get operation tables and store it in a list of dataframes
        pass

    def get_resource_table(self):
        # this will get resource tables and store it in a list of dataframes
        pass
