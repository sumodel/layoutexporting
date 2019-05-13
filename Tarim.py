# -*- coding: utf-8 -*-
from __future__ import print_function
import multiprocessing
from src import *
import gc

gc.enable()
multiprocessing.freeze_support()

if __name__ == '__main__':
    arcpy.env.overwriteOutput = True
    process_folder = r"C:\Users\knn\Desktop\_ARAS_FOR_TARIM_PAFTA_RISK"
    exemption = [

        # "SULTANDAGI_Derinlik_Tehlike.mxd"           ## 10 min
        "TARIM_TARIM_PAFTA_RISK.mxd",
        "TARIM_TARIM_PAFTA_RISK_V_0.1.mxd",
        "TARIM_TARIM_PAFTA_RISK_V_0.2.mxd",
        "TARIM_TARIM_PAFTA_RISK_V_0.3.mxd",
        "TARIM_TARIM_PAFTA_RISK_V_0.4.mxd",
    ]
    start = datetime.datetime.now()
    print(start)
    mxdfiles = list([])
    for dirpath, subdirs, files in os.walk(process_folder):
        for x in files:
            if x.endswith(".mxd") and x not in exemption:
                print(x)
                p = multiprocessing.Process(target=routine.process_it, args=(x, dirpath))
                p.start()
                p.join()
                print(x, "Duration", datetime.datetime.now() - start)
    print(datetime.datetime.now() - start)
