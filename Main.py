# -*- coding: utf-8 -*-
"""
Created on Fri Oct  8 12:13:32 2021

@author: dien.phan
"""

from ClassMC_Eff_Yield_Sum import Yield
from ClassMC_Eff_Yield_Sum import Eff
    

while True:
    run = input("Select running Type ( Yield Summary:1 / Machine efficiency:2 ) : ")
    if (run == "1") or (run == "2"):
        break
    else:
        print("[ Oops! That was no valid running type. Please, try again... ]") 
            
if run == "1":
    print("[ Yield  Summary start --> ]")
    r = Yield()
    r.YieldSummaryMain()
if run == "2":
    r = Eff()
    r.MachineEfficiencyMain()

print("[ DONE !!! (~.~!) ]")