# -*- coding: utf-8 -*-
from __future__ import print_function
import multiprocessing
from src import *
import gc
import sys

import time


if sys.argv[1] is not None:
    frame_no = sys.argv[1]
else:
    frame_no = "all"

print(frame_no)





gc.enable()
multiprocessing.freeze_support()

if __name__ == '__main__':
    arcpy.env.overwriteOutput = True
    process_folder = r"C:\Users\DHI\Desktop\aras_tarim_risk\_ARAS_FOR_TARIM_PAFTA_RISK\_ARAS_FOR_TARIM_PAFTA_RISK"
    exemption = [

        "TARIM_TARIM_PAFTA_RISK_V_0.5.mxd",
    ]
    start = datetime.datetime.now()
    print(start)
    mxdfiles = list([])
    for dirpath, subdirs, files in os.walk(process_folder):
        for x in files:
            if x.endswith(".mxd") and x not in exemption:
                print(x)
                p = multiprocessing.Process(target=routine.process_it, args=(x, dirpath, frame_no))
                p.start()
                p.join()
                print(x, "Duration", datetime.datetime.now() - start)
    print(datetime.datetime.now() - start)
