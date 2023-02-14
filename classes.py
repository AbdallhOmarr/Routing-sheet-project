
# This will contain all requrired product


# imports
import glob
import math
import os
import pandas as pd
import xlwings as xw
from collections import Counter


class Product:
    def __init__(self, code, description, main_category, sub_category, minor_category, item_type, weight, length, width, thickness, comp_qty, status, raw_material=None, locator=None, paint_qty=0, thinner_qty=0, galv_qty=0, welding_qty=0) -> None:

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

        # Paint data
        self.paint_qty = paint_qty
        self.thinner_qty = thinner_qty
        self.galv_qty = galv_qty

        # welding data
        self.welding_qty = welding_qty

        # manufacturing related attributes
        self.comp_qty = comp_qty

        # status attributes
        self.status = status

        # raw material
        self.raw_material = raw_material

        # product will have processes
        self.lst_of_processes = []

        self.std_route = False

    def get_product_vector(self,):
        # this will return product vector to calculate process cycle time
        product_vector = [self.main_category, self.sub_category, self.minor_category,
                          self.length, self.width, self.thickness, self.weight, self.comp_qty,
                          self.paint_qty, self.thinner_qty, self.galv_qty, self.welding_qty]
        product_vector = [x if x != None else 0 for x in product_vector]
        df = pd.DataFrame({'product_vector': product_vector})

        # Replace NaN values with 0
        df.fillna(0, inplace=True)

        # Convert the DataFrame back to a list
        product_vector = df['product_vector'].tolist()

        return product_vector

    def assign_process(self):
        # this will add a process obj to the lst of processes

        for i, v in self.route_processed.iterrows():
            # get process code
            department = v["department"]
            process = v["process"]
            machine = v["machine"]
            try:
                no_of_cuts = v["no of cuts"]
            except:
                no_of_cuts = 1
            op_seq = v["Op Seq"]
            res_seq = 10
            dept_data = StaticData().get_from_dept(department)
            process_data = StaticData().get_from_process(
                dept_data["id"].to_list()[0], process)
            if machine != "NaN":
                machine_data = StaticData().get_from_machine(
                    process_data["id"].to_list()[0], machine)

            labor_data = StaticData().get_from_labor(
                process_data["id"].to_list()[0])

            process = Process(process_data["code"].to_list()[
                              0], op_seq, no_of_cuts)
            process.assign_department(dept_data["code"].to_list()[0])
            if machine != "NaN":
                process.assign_machine(
                    machine_data["code"].to_list()[0], 1, res_seq)
                res_seq += 10

            for il, vl in labor_data.iterrows():
                process.assign_labor(vl["code"], 1, res_seq)
                res_seq += 10

            process.calc_rate(self.get_product_vector())
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

            if len(process.machines) >= 1:
                for machine in process.machines:
                    resource_data = {
                        "Part Code": self.code,
                        "Description": self.description,
                        "Operation Sequence": process.op_seq,
                        "Resource Sequence": machine.res_seq,
                        "Resource Code": machine.code,
                        "Assigned Units": machine.no_of_resource,
                        "Inverse": machine.rate,

                    }
                    lst_of_resource_data.append(resource_data)

            for labor in process.labors:
                resource_data = {
                    "Part Code": self.code,
                    "Description": self.description,
                    "Operation Sequence": process.op_seq,
                    "Resource Sequence": labor.res_seq,
                    "Resource Code": labor.code,
                    "Assigned Units": labor.no_of_resource,
                    "Inverse": labor.rate,

                }
                lst_of_resource_data.append(resource_data)

        return lst_of_resource_data

    def check_copy_route(self, df):

        self.copy_route = df["copy route"].to_list()
        if len(self.copy_route) >= 1:
            self.copy_route = self.copy_route[0]
        else:
            self.copy_route = None

    def get_route(self, df):
        if self.std_route == False:
            self.route = df.set_index("item code").stack().reset_index()
        else:
            self.std_routing_df = self.get_std_route(
                df["std route"].to_list()[0])
            merged = pd.merge(df, self.std_routing_df,
                              on="std route", how="left", suffixes=("_x", ""))
            merged.dropna(inplace=True, axis=1)
            self.route = merged.set_index("item code").stack().reset_index()

            # merge std route df with route df

        self.get_route_json()

    def get_route_json(self):
        self.route_processed = []
        route_dict = {
            "department": "dept1",
            "process": "process1",
            "machine": "machine1",
            "no of cuts": "no1"
        }
        for i, v in self.route.iterrows():
            route = {}
            if "dept" in v["level_1"]:
                route["department"] = v[0]
                route["Op Seq"] = int(v["level_1"][-1])*10

            if "process" in v["level_1"]:
                route["process"] = v[0]
                route["Op Seq"] = int(v["level_1"][-1])*10

            if "machine" in v["level_1"]:
                route["machine"] = v[0]
                route["Op Seq"] = int(v["level_1"][-1])*10

            if "no" in v["level_1"]:
                route["no of cuts"] = v[0]
                route["Op Seq"] = int(v["level_1"][-1])*10

            self.route_processed.append(route)

        self.route_processed = pd.DataFrame(self.route_processed)
        self.route_processed = self.route_processed.dropna(subset=["Op Seq"])

        grouped = self.route_processed.groupby("Op Seq")

        # Aggregate the data in each group as desired
        aggregated = grouped.agg(lambda x: x.value_counts(
        ).index[0] if x.notnull().any() else "NaN").reset_index()
        self.route_processed = aggregated.copy()

    def get_std_route(self, id):
        route = StaticData().get_from_std_routing(id)
        return route


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
        parts_code_start = ["622", "422", "322", "522"]
        lst_of_products = []
        parent_added = False
        for i, v in self.bom_df.iterrows():
            if parent_added == False:
                parent = Product(v["Top Parent"], v["Parent Description"], None, None, None, None, self.bom_df["Calc Unit Weight"].sum(
                ), self.bom_df["Comp Unit Length"].max(), self.bom_df["Comp Unit Width"].max(), None, None, v["Parent Item Status"])
                for ix, vx in self.bom_df.iterrows():
                    if vx["Comp Sub Category"] == "سلك اللحام":
                        parent.welding_qty += float(vx["Comp Qty"])
                    elif vx["Comp Sub Category"] == "البويات":
                        parent.paint_qty += float(vx["Comp Qty"])
                    elif vx["Comp Sub Category"] == "تنر":
                        parent.thinner_qty += float(vx["Comp Qty"])
                    elif vx["Comp Sub Category"] == "جلفنة":  # I should recheck this condition
                        parent.galv_qty += float(vx["Comp Qty"])
                self.top_parent = parent.code
                lst_of_products.append(parent)
                parent_added = True
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
        self.data = StaticData().get_from_dept_by_code(self.code)

    def get_wip(self):
        # get sheet for each dept wip
        # Filter on for current dept
        # return wip value
        return self.data["wip"].to_list()[0]


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
        self.machines = []
        self.labors = []

    def get_process_factors(self):
        # get factors to calc process cycle time
        # factors sheet is a path for excel

        if len(self.machines) >= 1:
            machine_code = self.machines[0].code
        else:
            machine_code = 0

        factors_df = StaticData().get_from_process_factors(self.code, machine_code)
        return factors_df

    def assign_department(self, dept_code):
        self.department = Department(dept_code)

        # after getting dept assign wip
        # value for each dept may be from Department class
        self.wip = self.department.get_wip()

    def assign_machine(self, machine_code, no_of_resource, res_seq):
        machine = Machine(machine_code, no_of_resource, res_seq)
        self.machines.append(machine)

    def assign_labor(self, labor_code, no_of_resource, res_seq):
        labor = Labor(labor_code, no_of_resource, res_seq)
        self.labors.append(labor)

    def check_no_of_resource(self, no=None):
        try:
            for machine in self.machines:
                machine.no_of_resource = Counter(self.machines)[machine]
        except:
            pass

        for labor in self.labors:
            labor.no_of_resource = Counter(self.labors)[labor]

    def calc_rate(self, product_vector):
        # in this function i will use product vector and factors_df to calc rate for this process
        # after calc rate
        # assign min order qty
        # value of multiple of 50 near the calc_rate

        # rate should be a dot product between product vector and process factors !?
        # for example the process factor for the saw process is as below

        max_rate = self.get_process_factors().values.tolist()[0][-2]
        min_rate = self.get_process_factors().values.tolist()[0][-1]

        process_vector = self.get_process_factors().values.tolist()[0][6:]
        product_vector = product_vector[3:]

        print(f"product_vector:{product_vector}")
        print(f"process vector: {process_vector}")
        self.rate = round(
            sum([a * b for a, b in zip(product_vector, process_vector)]), 2)

        print(f"rate :{self.rate}")
        self.check_no_of_resource()
        if self.no_of_cuts != "NaN":
            self.rate = self.rate / self.no_of_cuts
        else:
            print("errorr")
            pass

        # if self.rate > max_rate:
        #     self.rate = max_rate
        # elif self.rate < min_rate:
        #     self.rate = min_rate

        self.min_order_qty = round(self.rate/50)*50
        # assign rate for machine and labor
        try:
            for machine in self.machines:
                machine.assign_rate(self.rate)
        except:
            pass

        for labor in self.labors:
            labor.assign_rate(self.rate)


class Routing:
    def __init__(self, products) -> None:
        # this class to interact with Excel and Bom and product
        self.products = products

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
        lst_of_wip_data = []
        for product in self.products:
            lst_of_wip_data.append(product.get_wip_data())
        wip_df = pd.DataFrame(lst_of_wip_data)
        return wip_df

    def get_operation_data(self):
        # this will get operation data for each product
        # aggregate data into a list or dataframe
        lst_of_operations = []
        for product in self.products:
            lst_of_operations += product.get_operation_data()

        operation_df = pd.DataFrame(lst_of_operations)
        return operation_df

    def get_resource_data(self):
        # this will get resource data for each product
        # aggregate data into a list or dataframe
        lst_of_resources = []
        for product in self.products:
            lst_of_resources += product.get_resource_data()

        resource_df = pd.DataFrame(lst_of_resources)
        return resource_df


class StaticData:
    def __init__(self):
        self.wb = xw.Book.caller()
        self.load_department_excel()
        self.load_process_excel()
        self.load_machines_excel()
        self.load_labors_excel()
        self.load_process_factors_excel()
        self.load_std_routing()

    def load_department_excel(self):
        self.dept_df = self.wb.sheets["department"].range("A1:c100").options(
            pd.DataFrame, expand='table', index=False).value
        self.dept_df.dropna(inplace=True)

    def load_process_excel(self):
        self.process_df = self.wb.sheets["operations"].range("A1:E300").options(
            pd.DataFrame, expand='table', index=False).value
        self.process_df.dropna(inplace=True)

    def load_machines_excel(self):
        self.machine_df = self.wb.sheets["machines"].range("A1:D300").options(
            pd.DataFrame, expand='table', index=False).value
        self.machine_df.dropna(inplace=True)

    def load_labors_excel(self):
        self.labor_df = self.wb.sheets["labors"].range("A1:D300").options(
            pd.DataFrame, expand='table', index=False).value
        self.labor_df.dropna(inplace=True)

    def load_process_factors_excel(self):
        self.process_factors_df = self.wb.sheets["rates"].range("A1:D300").options(
            pd.DataFrame, expand='table', index=False).value
        self.process_factors_df.dropna(how="all", inplace=True)

    def load_std_routing(self):
        self.std_routing_df = self.wb.sheets["std routing"].range("A1:AY10").options(
            pd.DataFrame, expand="table", index=False
        ).value

    def get_from_dept_by_code(self, code):
        filtered_data = self.dept_df[self.dept_df["code"] == code]
        return filtered_data

    def get_from_dept(self, dept_desc):
        filtered_data = self.dept_df[self.dept_df["description"] == dept_desc]
        return filtered_data

    def get_from_process(self, dept_id, operation_desc):
        filtered_data = self.process_df[(self.process_df["description"] == operation_desc) & (
            self.process_df["department id"] == dept_id)]
        return filtered_data

    def get_from_machine(self, process_id, machine_desc):
        filtered_data = self.machine_df[(self.machine_df["description"] == machine_desc) & (
            self.machine_df["operation id"] == process_id)]
        return filtered_data

    def get_from_labor(self, process_id):
        filtered_data = self.labor_df[self.labor_df["operation id"] == process_id]
        return filtered_data

    def get_from_process_factors(self, process_code, machine_code):
        filtered_data = self.process_factors_df[(self.process_factors_df["process code"] == process_code) & (
            self.process_factors_df["machine code"] == machine_code)]
        return filtered_data

    def get_from_std_routing(self, id):
        filtered_data = self.std_routing_df[self.std_routing_df["std route"] == id]
        filtered_data.dropna(inplace=True, axis=1)

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
        self.cwd = os.path.realpath(
            os.path.join(os.getcwd(), os.path.dirname(__file__)))

    def get_parent_items(self, main_df):
        # load items from xlsx
        main_df = main_df.dropna()
        self.parent_items = []
        for i, v in main_df.iterrows():
            self.parent_items.append(v["Items Code"])

        print(f"parents:{self.parent_items}")
        print("-"*40)

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
