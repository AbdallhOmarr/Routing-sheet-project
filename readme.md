# User Story

- We want an Excel sheet to assign manufacturing processes on products to be manufactured, that has the following functions:
  1. load bill of material
  2. initialize product vector from bill of material which is a vector containing physical attributes of product
  3. assign factory, process, machine, and labors
  4. calculate the time of the manufacturing process
  5. transform data to suitable format to be uploaded on oracle production

# Objects Identified

1. Entity Objects:

   1. Bill Of Material
   2. Product
   3. Factory
   4. Manufacturing Process
   5. Machine
   6. Labor

2. Boundary Objects

   1. Product vector initialize
   2. Route data Format

3. Control Objects
   1. Load BOM
   2. Assign factory, process, machine, labor on product
   3. Calculate the time for manufacturing process
   4. transform data to upload data to Oracle production

.-

# processes rates to be changed

- Cut on Disk equation

# to do

- Add each paint qty, weld qty to each assembly by analysing the product tree and know which assembly is this raw mat belongs to
- Add up to 5 sheets
- Add set of standard routing
- Add send mail func
- Add dataloader func
- controller on routing to check max and min rate permitted

# Doing

- Add the ability to copy routing from another code in the workbook

# Done

- Add process factors algorithm to calculate rate
