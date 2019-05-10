# -*- coding: utf-8 -*-
"""

Danger and Depth Exporting
CreateDate = 10.11.2018
Author : Kenan Bolat
Version : 0.0.4
ModifyDate = 18.11.2018

"""

import arcpy
import re
import os
import datetime
import xlrd

import glob
import gc
import sys
import math
gc.enable()


# region Methods
def layer_visibility(map_document, layer_name, visible=True):
    for layer in arcpy.mapping.ListLayers(map_document, layer_name):
        layer.visible = visible
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()



def change_all_layers(map_document, enabled=False):
    layer_list = [u"Q100", u"Q10", u"Q50", u"Tehlike Haritaları", u"Derinlik Haritaları", u"Risk_Data"]
    for layer in layer_list:
        layer_visibility(map_document, layer, visible=enabled)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def enable_dimension_layers(map_document, enabled=True):
    layer_list = [u"1D", u"2D"]
    for layer in layer_list:
        layer_visibility(map_document, layer, visible=enabled)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_legend(map_document, flag_="default"):
    # flag = depth (X = 677, Y = 347, W =42, H  = 73 )
    # flag = danger (X = 677, Y = 318, W =42, H = 73 )
    # flag = default  (X = 850 )

    legend_type = "LEGEND_ELEMENT"
    legend_name = "LegendDepthAndDanger"
    if flag_ == "default":
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 850
            change_frame(map_document, flag_)
            # elm.elementPositionY = 325
    elif flag_ == "depth":
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 677
            elm.elementPositionY = 347
            change_frame(map_document, flag_)
    elif flag_ == "danger":
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 677
            elm.elementPositionY = 318
            change_frame(map_document, flag_)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_frame(map_document, flag_="default"):
    # flag = depth (X = 670, Y = 212, W =167, H  = 137 )
    # flag = danger (X = 670, Y = 236, W =167, H = 87 )
    # flag = default  (X = 850 )
    legend_type = "GRAPHIC_ELEMENT"
    frame_depth = "Frame_Depth"
    frame_danger = "Frame_Danger"
    if flag_ == "default":

        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, frame_danger):
            elm.elementPositionX = 850
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, frame_depth):
            elm.elementPositionX = 900
            # elm.elementPositionY = 325
    elif flag_ == "depth":
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, frame_danger):
            elm.elementPositionX = 850
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, frame_depth):
            elm.elementPositionX = 670
            # elm.elementPositionY = 210
            elm.elementPositionY = 213
    elif flag_ == "danger":
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, frame_danger):
            elm.elementPositionX = 670
            elm.elementPositionY = 236
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, frame_depth):
            elm.elementPositionX = 900
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_subtype(map_document, wsheet, subtype_=u'a', flag_="default"):
    if flag_ == "default":
        change_all_layers(map_document)
    elif flag_ == 'depth':
        change_all_layers(map_document)
        for layer in arcpy.mapping.ListLayers(map_document):
            if layer.longName == subtype_ + "\\" + u"Derinlik Haritaları":
                layer_visibility(map_document, layer, visible=True)
                layer_visibility(map_document, subtype_, visible=True)
                change_legend(map_document, flag_)
                update_text(map_document, subtype_)
                update_pafta_value(map_document, flag_, subtype_, wsheet)
    elif flag_ == 'danger':
        change_all_layers(map_document)
        for layer in arcpy.mapping.ListLayers(map_document):
            if layer.longName == subtype_ + "\\" + u"Tehlike Haritaları":
                layer_visibility(map_document, layer, visible=True)
                layer_visibility(map_document, subtype_, visible=True)
                change_legend(map_document, flag_)
                update_text(map_document, subtype_)
                update_pafta_value(map_document, flag_, subtype_, wsheet)


def update_text(map_document, subtype_):
    legend_elemen_list = [u"discharge_label_upper", u"Antet_Subtype"]
    for leg in legend_elemen_list:
        for el in arcpy.maupdate_textpping.ListLayoutElements(map_document, "TEXT_ELEMENT", leg):
            el.text = re.sub("(?<=<SUB>)(.*?)(?=<\/SUB>)", subtype_[1:], el.text)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def update_pafta_value(map_document, flag, subtype_, wsheet):
    legend_elemen_list = [u"Pafta_Name", u"Q_VAL_1", u"Q_VAL_2", u"Q_VAL_3", u"Q_VAL_4", u"Q_VAL_5"]
    input__sup = get_input_values(wsheet, 1)
    for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Antet_Subtype_sup"):
        el.text = input__sup[2]
    if flag == 'danger':
        for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Antet_Subtype"):
            el.text = re.sub("(?<=^)(.*?)(?=\()", u"TAŞKIN TEHLİKE HARİTASI ", el.text)
        if subtype_ == "Q50":
            input_ = get_input_values(wsheet, 5)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "

        if subtype_ == "Q100":
            input_ = get_input_values(wsheet, 6)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
        if subtype_ == "Q500":
            input_ = get_input_values(wsheet, 6)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
        if subtype_ == "Q10":
            input_ = get_input_values(wsheet, 4)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
    if flag == 'depth':
        for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Antet_Subtype"):
            el.text = re.sub("(?<=^)(.*?)(?=\()", u"TAŞKIN SU DERİNLİĞİ HARİTASI", el.text)
        if subtype_ == "Q50":
            input_ = get_input_values(wsheet, 2)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
        if subtype_ == "Q100":
            input_ = get_input_values(wsheet, 3)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
        if subtype_ == "Q500":
            input_ = get_input_values(wsheet, 1)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
        if subtype_ == "Q10":
            input_ = get_input_values(wsheet, 1)
            for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Pafta_Name"):
                el.text = input_[1]
            for en, agi in enumerate(legend_elemen_list[1:]):
                # print en + 1, agi
                for el_agi in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", agi):
                    try:
                        el_agi.text = str(input_[4]).split(",")[en].replace(" ", "")
                    except:
                        el_agi.text = u" "
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def get_input_values(wsheet, column_index):
    data_ = []
    for rows in range(wsheet.nrows):
        data_.append(wsheet.cell_value(rows, column_index))
    return data_


def change_subtype_layer_names(map_document):
    for lyr in arcpy.mapping.ListLayers(map_document, u"(Q*"):
        # print lyr.longName, lyr.name
        if lyr.longName.split("\\")[0] == ('Q500'):
            lyr.name = re.sub("(?<=<SUB>)(.+?)(?=<\/SUB>)", "500", lyr.name)
        if lyr.longName.split("\\")[0] == ('Q100'):
            lyr.name = re.sub("(?<=<SUB>)(.+?)(?=<\/SUB>)", "100", lyr.name)
        if lyr.longName.split("\\")[0] == ('Q50'):
            lyr.name = re.sub("(?<=<SUB>)(.+?)(?=<\/SUB>)", "50", lyr.name)
        if lyr.longName.split("\\")[0] == ('Q10'):
            lyr.name = re.sub("(?<=<SUB>)(.+?)(?=<\/SUB>)", "10", lyr.name)
    for l in arcpy.mapping.ListLayers(map_document, "*D"):
        l.visible = True
    for l in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları *"):
        l.visible = True

    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_upper_antet(map_document, peak_count, show=True):
    for l in arcpy.mapping.ListLayoutElements(map_document, "GRAPHIC_ELEMENT", "Antet_group*"):
        l.elementPositionX = 871.2805
        arcpy.RefreshActiveView()
    if show:
        for l in arcpy.mapping.ListLayoutElements(map_document, "GRAPHIC_ELEMENT", "Antet_group_" + str(peak_count)):
            l.elementPositionX = 670.8862
            l.elementPositionY = 352.2442
        arcpy.RefreshActiveView()
    else:
        for l in arcpy.mapping.ListLayoutElements(map_document, "GRAPHIC_ELEMENT", "Antet_group*"):
            l.elementPositionX = 871.2805
            arcpy.RefreshActiveView()


def update_peak_values_text(map_document, peak_values, peak_texts):
    for en_val, val_ in enumerate(peak_values):
        for el_val in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Q_VAL_" + str(en_val+1)):
            el_val.text = val_
    for en, txt in enumerate(peak_texts):
        for el_text in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Q_VAL_TEXT_" + str(en+1)):
            el_text.text = txt
    arcpy.RefreshActiveView()


def update_text(map_document, subtype_):
    legend_elemen_list = [u"discharge_label_upper", u"Antet_Subtype"]
    for leg in legend_elemen_list:
        for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", leg):
            el.text = re.sub("(?<=<SUB>)(.*?)(?=<\/SUB>)", subtype_[1:], el.text)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()

def process_it(x_,  dirpath_):
    location_of_mxd = os.path.join(dirpath_, x_)
    path_of_process = os.path.dirname(location_of_mxd)
    excel_folder_name = os.path.abspath(os.path.join(path_of_process, "..", "EXCEL"))

    output_pdf_path = os.path.abspath(os.path.join(path_of_process, "..", "PDF"))
    mxd = arcpy.mapping.MapDocument(location_of_mxd)
    # mxd = arcpy.mapping.MapDocument("CURRENT")
    # temp_mxd = arcpy.mapping.MapDocument(location_of_mxd)

    df_data = arcpy.mapping.ListDataFrames(mxd, "Data")[0]
    for lyr_data in arcpy.mapping.ListLayers(mxd, "25000*", df_data):
        pass
    arcpy.SelectLayerByAttribute_management(lyr_data, "CLEAR_SELECTION")
    frame_no_list = [row[0] for row in arcpy.da.SearchCursor(lyr_data, "Frame_No")]

    for frame in frame_no_list:
        print frame
        excel_file_name = os.path.join(excel_folder_name, "FR"+str(int(frame))+".xlsx")
        excel_file = os.path.join(excel_folder_name, excel_file_name)
        workbook = xlrd.open_workbook(excel_file)
        worksheet = workbook.sheet_by_index(0)
        layer_visibility(mxd, "Basemap", False)
        df = arcpy.mapping.ListDataFrames(mxd, "LegendUR*")[0]
        for lyr in arcpy.mapping.ListLayers(mxd, "25000*", df):
            pass

        arcpy.SelectLayerByAttribute_management(lyr, "NEW_SELECTION", "Frame_No = " + str(int(frame)))
        df_data.extent = lyr.getSelectedExtent()
        df_data.scale = 25000
        change_subtype_layer_names(mxd)
        subtype_list = [u"Q50", u"Q100", u"Q10"]
        flag_list = [u'danger', u'depth']
        for flag in flag_list:
            for subtype in subtype_list:
                change_subtype(mxd, worksheet, subtype, flag)
                if flag == u'depth':
                    tag = u'Derinlik'
                elif flag == u'danger':
                    tag = u'Tehlike'
                try:

                    if flag == u'depth':
                        arti = 0
                    else:
                        arti = 3
                    if subtype == u"Q10":
                        flag_count = 1 + arti
                    elif subtype == u"Q50":
                        flag_count = 2 + arti
                    elif subtype == u"Q100":
                        flag_count = 3 + arti

                    # p_t = get_input_values(worksheet, 0)[4].split(";")[1].split(",")
                    # p_v = str(get_input_values(worksheet, flag_count)[4]).split(",")
                    # update_peak_values_text(mxd, p_v, p_t)
                    # change_upper_antet(mxd, len(p_v))
                    update_text(mxd, subtype)
                    arcpy.RefreshTOC()
                    arcpy.RefreshActiveView()
                    # print "Exporting ... ", flag, ":", subtype
                    arcpy.mapping.ExportToPDF(mxd, os.path.join(output_pdf_path, "FR" + str(int(frame)) + "_" +
                                                                x_.split(".")[
                                                                    0] + "_" + tag + "_" + subtype + ".pdf"))
                    # print "Exporting ... \t\t", flag, ":\t\t\t", subtype, "Exported Successfully"
                    print '%12s %12s %12s %12s' % ('Exporting', flag, subtype, "Successful")

                except:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(exc_type, fname, exc_tb.tb_lineno)
                    print "!!HATA ... \t\t", flag, ":\t\t\t", subtype, "Exporting Unsuccsesful"
    del mxd

# endregion


# process_folder = r"C:\Users\knn\Desktop\TARIM_PAFTA_SET\TARIM_PAFTA_SET"
process_folder = r"C:\Users\knn\Desktop\_ARAS_FOR_TARIM_PAFTA_RISK"
# process_folder = r"I:\temp\sumodel\BM_AC_BA_BU\AC\A   C"
#

exemption = [

# "AFYON_Derinlik_Tehlike.mxd",               ## 10 min
# "AKSEHIR_Derinlik_Tehlike.mxd",             ## 10 min
# "CAY_Derinlik_Tehlike.mxd",                 ## 10 min
# "DERECINE_Derinlik_Tehlike.mxd",            ## 10 min
# "GAZLIGOL_Derinlik_Tehlike.mxd",            ## 10 min
# "ISCEHISAR_Derinlik_Tehlike.mxd",           ## 10 min
# "SEYITLER3_Derinlik_Tehlike.mxd",           ## 10 min
# "SINANPASA_Derinlik_Tehlike.mxd",           ## 10 min
# "SUHUT_Derinlik_Tehlike.mxd",               ## 10 min
# "SULTANDAGI_Derinlik_Tehlike.mxd"           ## 10 min

]


import multiprocessing
multiprocessing.freeze_support()
if __name__ == '__main__':
    arcpy.env.overwriteOutput = True

    start = datetime.datetime.now()
    print start
    mxdfiles = list([])

    for dirpath, subdirs, files in os.walk(process_folder):
        for x in files:
            if x.endswith(".mxd") and x not in exemption:
                # if x.endswith(".mxd"):

                print x
                p = multiprocessing.Process(target=process_it, args=(x, dirpath))
                p.start()
                p.join()
                #
                # del mxd
                # del excel_file_name
                # del excel_file
                # del worksheet
                # del workbook
                print x, "Duration", datetime.datetime.now() - start

    print datetime.datetime.now() - start
