
# This will contain all requrired product


# imports
import pandas as pd


class Product:
    def __init__(self, code, description, main_category, sub_category, minor_category, item_type, weight, length, width, thickness, comp_qty, make_buy, status, raw_material=None, locator=None) -> None:

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
        self.make_buy = make_buy
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

    def assign_process(self, code, op_seq):
        # this will add a process obj to the lst of processes
        process = Process(code, op_seq)

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
    def __init__(self, bom_file) -> None:
        # in this class i will get bom excel sheet and extract products and raw material from it.

        # bom_file is the file path for bom
        self.bom_file = bom_file

    def get_bom_df(self):
        # - convert bom_file to dataframe
        self.bom_df = pd.read_excel(self.bom_file)

    def get_lst_of_products(self):
        # - each line will contain a data of a product

        # loop on the dataframe
        # inialize products
        # append to lst_of_products
        # return list for products
        lst_of_products = []

    def get_route_df(self):
        # in this function i will return data to enable user to assign factory, process, machine, no of labors
        pass


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
    def __init__(self, code, op_seq, min_order_qty=None, wip=None) -> None:
        self.code = code
        self.op_seq = op_seq
        self.min_order_qty = min_order_qty
        self.wip = wip

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
        self.rate = "rate"
        self.min_order_qty = "min order qty"
        # assign rate for machine and labor
        self.machine.assign_rate(self.rate)
        self.labor.assign_rate(self.rate)


class Routing:
    def __init__(self) -> None:
        # this class to transform data into suitable format for oracle
        pass
