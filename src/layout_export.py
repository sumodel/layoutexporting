# -*- coding: utf-8 -*-
# region Modules

import arcpy
import re
import os
import datetime
import math
import xlrd
import glob
# endregion

arcpy.env.overwriteOutput = True
start = datetime.datetime.now()
print start

#TODO add to routine class


settings_flag = {
    # "process_folder": r"C:\Users\DHI\Desktop\_YAS_RISK\BM_2\1a\BANAZ",
    "process_folder": r"C:\Users\DHI\Desktop\_YAS_RISK\_OLUDENIZ_PAFTA_SET",
    "mahalle": r"C:\Users\DHI\Desktop\_YAS_RISK\ORTAK_SHPyeniler\Mahalle_v14_v2_tm30.shp",
    # "Ekonomik_Zarar_Colormap": r"C:\Users\DHI\Desktop\BM_1\_ORTAK_SHP\Ekonomik_Zarar_Colormap.lyr",
    "Ekonomik_Zarar_Colormap": r"C:\Users\DHI\Desktop\Kenan_Taslaklar\Template\Risk_Template\SHP\Ekonomik_Zarar_Colormap.lyr",
    # "Risk_Colormap": r"C:\Users\DHI\Desktop\BM_1\_ORTAK_SHP\Risk_Colormap.lyr",
    "Risk_Colormap": r"C:\Users\DHI\Desktop\Kenan_Taslaklar\Template\Risk_Template\SHP\Risk_Colormap.lyr",
    "nufus": "C:\Users\DHI\Desktop\Kenan_Taslaklar\Template\Risk_Template\SHP\Nufus.lyr",
    "Cevre": False,
    "Aktivite": False,
    "Stratejik": False,
    "Risk": False,
    "Zarar": True,
    "Nufus": False,
    "Basemap": True,

}
process_folder = settings_flag["process_folder"]


def layer_visibility(map_document, layer_name, visible=True):
    for layer in arcpy.mapping.ListLayers(map_document, layer_name):
        layer.visible = visible
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_all_layers(map_document, enabled=False):
    layer_list = ["Risk_Zarar", "Genel", "Aktivite", "Stratejik", "Cevre", "Kesisim", "BUILDING_RISK_v4"]
    for layer in layer_list:
        layer_visibility(map_document, layer, visible=enabled)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_legend_gosterim(map_document, enabled=True):
    legend_type = "LEGEND_ELEMENT"
    legend_name = "LegendGosterim"
    layer_risk_zarar = "Risk_Zarar"
    if enabled:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 681
            elm.elementPositionY = 325
            elm.title = u"Gösterim"
        layer_visibility(map_document, layer_risk_zarar)
    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 857
            elm.elementPositionY = 325
        layer_visibility(map_document, layer_risk_zarar, visible=False)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_legend_nufus(map_document, enabled=True):
    legend_type = "GRAPHIC_ELEMENT"
    legend_name = "Nufus_Legend"
    if enabled:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 678
            elm.elementPositionY = 211
            layer_visibility(map_document, "Nufus", visible=True)
    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 857
            elm.elementPositionY = 325
        layer_visibility(map_document, "Etkilenen_Nufus", visible=False)
        layer_visibility(map_document, "Nufus", visible=False)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_legend_cevre_aktivite_stratejik(map_document, subtype, tag=u" ", enabled=True):
    legend_type = "LEGEND_ELEMENT"
    legend_name = "ThirdLegend"
    disabling_layers = ["Genel", "Aktivite", "Stratejik", "Cevre", "Risk", "Risk_Zarar", "Kesisim"]
    for l in disabling_layers:
        layer_visibility(map_document, l, visible=False)
    if enabled:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 678
            elm.elementPositionY = 322
            # elm.elementHeight = 113
            # elm.elementWidth = 149
            elm.title = tag
        for l in ["Genel", "Kesisim", subtype]:
            layer_visibility(map_document, l)
    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 890
            elm.elementPositionY = 230
            # elm.elementHeight = 20
            # elm.elementWidth = 20
        for l in disabling_layers:
            layer_visibility(map_document, l, visible=False)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def disable_old_antets(mxd_):
    old_antet_list = ["q_discharge_peak_table_title",
                      "q_discharge_peak_address",
                      "q_discharge_Peak_Value",
                      "q_discharge_peak_label"
                      ]
    for l in arcpy.mapping.ListLayoutElements(mxd_, "TEXT_ELEMENT"):
        if l.name in old_antet_list:
            # print l.name, "will be deleted"
            l.text = u" "
    del l
    for l in arcpy.mapping.ListLayoutElements(mxd_, "PICTURE_ELEMENT", "Risk_Antet_Transparent"):
        if l.name == "Risk_Antet_Transparent":
            l.elementPositionX = 1000
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_legend_zarar(map_document, enabled=True):
    legend_type = "LEGEND_ELEMENT"
    legend_name = "LegendZarar"
    layer_risk_zarar = "Risk_Zarar"
    if enabled:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 678
            elm.elementPositionY = 282
            # elm.elementHeight = 69
            # elm.elementWidth = 52
            elm.title = u"Ekonomik Zarar (TL)"
        layer_visibility(map_document, layer_risk_zarar)
    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 1100
            elm.elementPositionY = 279
        layer_visibility(map_document, layer_risk_zarar, visible=False)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_data_frame_taskin_alani(map_document, enabled=True):
    legend_type = "DATAFRAME_ELEMENT"
    legend_name = "TaskinAlani"
    if enabled:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            # elm.elementPositionX = 681
            # elm.elementPositionY = 527
            # elm.elementWidth = 140
            elm.elementHeight = 155
            data_frame = arcpy.mapping.ListDataFrames(map_document, legend_name)[0]
            map_layers = arcpy.mapping.ListLayers(map_document, "dere_model_*")
            layer = map_layers[0]
            extent = layer.getExtent(True)
            data_frame.extent = extent

    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            # elm.elementPositionX = 681
            # elm.elementPositionY = 527
            # elm.elementWidth = 140
            elm.elementHeight = 100
            data_frame = arcpy.mapping.ListDataFrames(map_document, legend_name)[0]
            map_layers = arcpy.mapping.ListLayers(map_document, "dere_model_*")
            layer = map_layers[0]
            extent = layer.getExtent(True)
            data_frame.extent = extent
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_legend_risk(map_document, enabled=True, mode="main"):
    legend_type = "PICTURE_ELEMENT"
    legend_name = "Risk_Legend"
    layer_risk_zarar = "Risk_Zarar"
    if enabled:
        if mode == "main":
            for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
                elm.elementPositionX = 678
                elm.elementPositionY = 219
                elm.elementHeight = 66
                # elm.elementWidth = 56

            layer_visibility(map_document, layer_risk_zarar)
        elif mode == "secondary":
            for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
                elm.elementPositionX = 722
                elm.elementPositionY = 326
                elm.elementHeight = 93
                # elm.elementWidth = 56

            layer_visibility(map_document, layer_risk_zarar, visible=False)
    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 849
            elm.elementPositionY = 229
            # elm.elementHeight = 93
            # elm.elementWidth = 56
        layer_visibility(map_document, layer_risk_zarar, visible=False)


def change_antet_risk_zarar(map_document, change_dict, enabled=True):
    legend_type = "PICTURE_ELEMENT"
    legend_name = "Risk_Antet_Transparent"
    if enabled:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 678
            elm.elementPositionY = 357
            # elm.elementHeight = 31
            # elm.elementWidth = 156
        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT"):
            for text in change_dict.keys():
                if elm.name == text:
                    elm.text = change_dict[text]
    else:
        for elm in arcpy.mapping.ListLayoutElements(map_document, legend_type, legend_name):
            elm.elementPositionX = 950
            elm.elementPositionY = 357
        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT"):
            if elm.name in change_dict.keys():
                elm.text = u" "


def remove_unwanted_layers_from_legends(mxd_, unwanted_list):
    for l in arcpy.mapping.ListLayoutElements(mxd_, "LEGEND_ELEMENT"):
        for lyrs in l.listLegendItemLayers():
            if lyrs.longName.find("Risk") != -1:
                l.removeItem(lyrs)
            if lyrs.name in unwanted_list:
                l.removeItem(lyrs)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def run_risk(map_document, temp_map_document, stats, mode="main", file_tag="risk_"):
    for LYR_1 in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları*"):
        LYR_1.visible = False
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
    for stat in stats:
        try:
            print "Initiate ", stat[0]
            to_be_change = ["SubLegend_Address", "SubtypeDischarge"]
            for LYR_2 in arcpy.mapping.ListLayers(map_document, "BUILDING_RISK_v4"):
                LYR_2.symbology.valueField = stat[2]
                LYR_2.symbology.reclassify()
                LYR_2.visible = True
                LYR_2.symbology.classBreakValues = return_break_values(temp_map_document, stat[2],
                                                                       "BUILDING_RISK_v4")
                LYR_2.definitionQuery = '"' + stat[2] + '" > 0 '
                arcpy.RefreshTOC()
                arcpy.RefreshActiveView()
            if mode == "main":
                for LYR_3 in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları*"):
                    if LYR_3.longName.find(stat[4]) != -1 or LYR_3.longName.find(stat[5]) != -1:
                        LYR_3.visible = True
                    else:
                        LYR_3.visible = False
                    arcpy.RefreshTOC()
                    arcpy.RefreshActiveView()
                for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT"):
                    if elm.name in to_be_change:
                        # print elm.name, elm.text
                        elm.text = re.sub("<SUB>(.|\n)*?</SUB>", stat[1], elm.text)
                    if elm.name == "q_discharge_Peak_Value":
                        elm.text = stat[3]
            else:
                for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT"):
                    if elm.name in to_be_change:
                        # print elm.name, elm.text
                        elm.text = re.sub("<SUB>(.|\n)*?</SUB>", stat[1], elm.text)
                pass

            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()
            if mode == "main":
                fname = os.path.join(output_path, file_tag + stat[0] + ".pdf")
                layer_visibility(map_document, "BUILDING_RISK_v4")
                for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "PaftaNo"):
                    elm.text = stat[8]
                for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubtypeDischarge"):
                    elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', stat[14], elm.text, re.MULTILINE)
                for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubLegend_Address"):
                    elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', stat[14], elm.text, re.MULTILINE)
                arcpy.RefreshTOC()
                arcpy.RefreshActiveView()
                p_t = get_input_values(worksheet, 0)[3].split(";")[1].split(",")
                if stat[0] == u"q50":
                    flag_count = 1
                elif stat[0] == u"q100":
                    flag_count = 2
                elif stat[0] == u"q500":
                    flag_count = 3

                p_v = str(get_input_values(worksheet, flag_count)[3]).split(",")
                update_peak_values_text(map_document, p_v, p_t)
                change_upper_antet(map_document, len(p_v))
                update_text(map_document, stat[0])
                change_oned_twod(map_document, stat[0][1:])
                disable_old_antets(map_document)

                colormap_temp_layers = ["risk_template", "ekonomik_zarar_template"]
                layer_visibility(map_document, colormap_temp_layers[0], visible=False)
                layer_visibility(map_document, colormap_temp_layers[1], visible=False)

                print "Exporting..", fname
                arcpy.mapping.ExportToPDF(map_document, fname)
                for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubLegend_Address"):
                    elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', u'TAŞKIN RİSK HARİTASI', elm.text, re.MULTILINE)
        except:
            print "Hata Var" , stat[0]


def get_intersection_values(map_document, query_list, intersect_layer):
    try:
        q__layer = "q__layer"
        i__layer = "i__layer"
        arcpy.env.workspace = "C:\\Users\\Hp\\Documents\\ArcGIS\\Default.gdb"
        arcpy.env.overwriteOutput = True

        for LYR in arcpy.mapping.ListLayers(map_document):
            if LYR.longName == intersect_layer:
                arcpy.MakeFeatureLayer_management(LYR, q__layer)
                break
        value = 0
        for row in query_list:
            for inner_lyr in arcpy.mapping.ListLayers(map_document):
                if inner_lyr.longName == row:
                    arcpy.MakeFeatureLayer_management(inner_lyr, i__layer)
            arcpy.SelectLayerByLocation_management(i__layer, "INTERSECT", q__layer, "", "NEW_SELECTION")
            value = value + len([row for row in arcpy.da.SearchCursor(i__layer, "*")])
            arcpy.SelectLayerByAttribute_management(i__layer, "CLEAR_SELECTION")
            arcpy.SelectLayerByAttribute_management(q__layer, "CLEAR_SELECTION")
    except BaseException as Be:
        pass
    return value


def clear_selection(map_document):
    for lyrs in arcpy.mapping.ListLayers(map_document):
        # noinspection PyBroadException
        try:
            arcpy.SelectLayerByAttribute_management(lyrs, "CLEAR_SELECTION")
        except Exception:
            continue


def return_break_values(map_document_t, field_name, building_layer):
    for lyr in arcpy.mapping.ListLayers(map_document_t, building_layer):
        lyr.symbology.valueField = field_name
        lyr.symbology.reclassify()
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
        return lyr.symbology.classBreakValues


def run_zarar(map_document, temp_map_document, stats, file_tag="Zarar_"):
    for LYR_1 in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları*"):
        LYR_1.visible = False
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
    del LYR_1
    for stat in stats:
        try:
            print "Initiate ", stat[0]
            tobe_changed = ["SubLegend_Address", "SubtypeDischarge"]
            for LYR_2 in arcpy.mapping.ListLayers(map_document, "BUILDING_RISK_v4"):
                field_fin = fieldname_check(map_document, u"BUILDING_RISK_v4", stat[6])
                LYR_2.symbology.valueField = field_fin
                LYR_2.symbology.reclassify()
                LYR_2.visible = True
                class_break_values = return_break_values(temp_map_document, field_fin, "BUILDING_RISK_v4")
                # class_break_values = LYR_2.symbology.classBreakValues
                zarar_multiplier = 1000
                LYR_2.definitionQuery = '"' + field_fin + '" > 0 '
                for en, value in enumerate(class_break_values):
                    if en > 0:
                        class_break_values[en] = int(
                            (math.ceil(float(value + zarar_multiplier) / zarar_multiplier)) * zarar_multiplier)
                    else:
                        class_break_values[0] = 1
                class_break_class = []
                for en, value in enumerate(class_break_values):
                    # print en, value
                    if en == 1:
                        class_break_class.extend([str(class_break_values[en - 1]) + "- " + str(value)])
                    if en > 1:
                        class_break_class.extend([str(class_break_values[en - 1] + 1) + "- " + str(value)])
                LYR_2.symbology.classBreakValues = class_break_values
                LYR_2.symbology.classBreakLabels = class_break_class
                LYR_2.symbology.reclassify()
            del LYR_2
            for LYR_3 in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları*"):
                if LYR_3.longName.find(stat[4]) != -1 or LYR_3.longName.find(stat[5]) != -1:
                    LYR_3.visible = True
                else:
                    LYR_3.visible = False
                arcpy.RefreshTOC()
                arcpy.RefreshActiveView()
            del LYR_3
            for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT"):
                if elm.name in tobe_changed:
                    # print elm.name, elm.text
                    elm.text = re.sub("<SUB>(.|\n)*?</SUB>", stat[1], elm.text)
                if elm.name == "q_discharge_Peak_Value":
                    elm.text = stat[3]
            for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "PaftaNo"):
                elm.text = stat[9]

            for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubtypeDischarge"):
                elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', stat[15], elm.text, re.MULTILINE)
            for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubLegend_Address"):
                elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', stat[15], elm.text, re.MULTILINE)
            fname = os.path.join(output_path, file_tag + stat[0] + ".pdf")
            p_t = get_input_values(worksheet, 0)[3].split(";")[1].split(",")
            if stat[0] == u"q50":
                flag_count = 1
            elif stat[0] == u"q100":
                flag_count = 2
            elif stat[0] == u"q500":
                flag_count = 3

            p_v = str(get_input_values(worksheet, flag_count)[3]).split(",")
            update_peak_values_text(map_document, p_v, p_t)
            change_upper_antet(map_document, len(p_v))
            update_text(map_document, stat[0])
            for ll in arcpy.mapping.ListLayers(mxd, u"Taşkın*"):
                ll.visible = False
            for ll in arcpy.mapping.ListLayers(mxd, u"Taşkın*(" + str(stat[6][1:]) + ")"):
                ll.visible = True

            layer_visibility(map_document, "Raster", True)
            layer_visibility(map_document, "*D", True)
            change_oned_twod(map_document, stat[0][1:])
            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()
            print "Exporting..", fname
            arcpy.mapping.ExportToPDF(map_document, fname)
            for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubLegend_Address"):
                elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', u'TAŞKIN RİSK HARİTASI', elm.text, re.MULTILINE)
        except:
            print "Hata var", stat[0]


def get_input_values(wsheet, column_index):
    data_ = []
    for rows in range(wsheet.nrows):
        data_.append(wsheet.cell_value(rows, column_index))
    return data_


def change_upper_antet(map_document, peak_count, show=True):
    if show:
        for l in arcpy.mapping.ListLayoutElements(map_document, "GRAPHIC_ELEMENT", "Antet_group_" + str(peak_count)):
            l.elementPositionX = 671.2805
            l.elementPositionY = 334.8558
        arcpy.RefreshActiveView()
    else:
        for l in arcpy.mapping.ListLayoutElements(map_document, "GRAPHIC_ELEMENT", "Antet_group*"):
            l.elementPositionX = 871.2805
            arcpy.RefreshActiveView()


def update_peak_values_text(map_document, peak_values, peak_texts):
    for en_val, val_ in enumerate(peak_values):
        for el_val in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Q_VAL_" + str(en_val + 1)):
            el_val.text = val_
    for en, txt in enumerate(peak_texts):
        for el_text in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "Q_VAL_TEXT_" + str(en + 1)):
            el_text.text = txt
    arcpy.RefreshActiveView()


def update_text(map_document, subtype_):
    legend_elemen_list = [u"discharge_label_upper", u"Antet_Subtype"]
    for leg in legend_elemen_list:
        for el in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", leg):
            el.text = re.sub("(?<=<SUB>)(.*?)(?=<\/SUB>)", subtype_[1:], el.text)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def change_oned_twod(map_document, subtype):
    layer_visibility(map_document, u"Taşkın Alanları Su Yükseklikleri (m)(50)", False)
    layer_visibility(map_document, u"Taşkın Alanları Su Yükseklikleri (m)(100)", False)
    layer_visibility(map_document, u"Taşkın Alanları Su Yükseklikleri (m)(500)", False)
    layer_visibility(map_document, u"Taşkın Alanları Su Yükseklikleri (m)(" + str(subtype) + ")", True)
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView()


def fieldname_check(map_document, layername, field_alias_name):
    temp_lyr_for_field = "temp_lyr_for_field"
    for flyr in arcpy.mapping.ListLayers(map_document, layername):
        if flyr.name.find(layername) != -1:
            print flyr.name
            break
    arcpy.MakeFeatureLayer_management(flyr, temp_lyr_for_field)
    return [row.name for row in arcpy.ListFields(temp_lyr_for_field) if row.aliasName == field_alias_name][0]


def update_nufus_counts(map_document, subtype):
    temp_lyr_for_sums = "temp_lyr_for_field_sums"
    mahalle = "mahalle"
    for ll_mah in arcpy.mapping.ListLayers(map_document, "Mahalle_Count"):
        if ll_mah.name == "Mahalle_Count":
            arcpy.MakeFeatureLayer_management(ll_mah, mahalle)
            break

    temp = []
    for ll in arcpy.mapping.ListLayers(map_document, "*nuf" + str(subtype[1:])):
        arcpy.MakeFeatureLayer_management(ll, temp_lyr_for_sums)
        t_adi = \
            list(set(
                [row[0] for row in arcpy.da.SearchCursor(temp_lyr_for_sums, ["adi_dum", "etnuf_" + str(subtype[1:])])]))[0]
        t_sum = sum([row[1] for row in arcpy.da.SearchCursor(temp_lyr_for_sums, ["adi_dum", "etnuf_" + str(subtype[1:])])])
        temp.append([t_adi, t_sum])
    for el in temp:
        with arcpy.da.UpdateCursor(mahalle, ["adi_dum", "Count"], "adi_dum = '" + el[0] + "'") as cursor:
            for row in cursor:
                if row[0] == el[0]:
                    cursor.updateRow(el)


def run_nufus(map_document, stats, file_tag="Nufus_"):
    for LYR_1 in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları*"):
        LYR_1.visible = False
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
    del LYR_1
    for stat in stats:
        print "Initiate ", stat[0]
        tobe_changed = ["SubLegend_Address", "SubtypeDischarge"]
        for LYR_3 in arcpy.mapping.ListLayers(map_document, u"Taşkın Alanları*"):
            if LYR_3.longName.find(stat[4]) != -1 or LYR_3.longName.find(stat[5]) != -1:
                LYR_3.visible = True
            else:
                LYR_3.visible = False
            arcpy.RefreshTOC()
            arcpy.RefreshActiveView()
        del LYR_3
        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT"):
            if elm.name in tobe_changed:
                # print elm.name, elm.text
                elm.text = re.sub("<SUB>(.|\n)*?</SUB>", stat[1], elm.text)
            if elm.name == "q_discharge_Peak_Value":
                elm.text = stat[3]
        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "PaftaNo"):
            elm.text = stat[10]

        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubtypeDischarge"):
            elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', stat[16], elm.text, re.MULTILINE)
        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubLegend_Address"):
            elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', stat[16], elm.text, re.MULTILINE)
        fname = os.path.join(output_path, file_tag + stat[0] + ".pdf")
        p_t = get_input_values(worksheet, 0)[3].split(";")[1].split(",")
        if stat[0] == u"q50":
            flag_count = 1
        elif stat[0] == u"q100":
            flag_count = 2
        elif stat[0] == u"q500":
            flag_count = 3
        update_nufus_counts(map_document, stat[0])

        p_v = str(get_input_values(worksheet, flag_count)[3]).split(",")
        update_peak_values_text(map_document, p_v, p_t)
        change_upper_antet(map_document, len(p_v))
        update_text(map_document, stat[0])
        colormap_temp_layers = ["risk_template", "ekonomik_zarar_template"]
        layer_visibility(map_document, colormap_temp_layers[0], visible=False)
        layer_visibility(map_document, colormap_temp_layers[1], visible=False)
        change_legend_risk(map_document, False)
        layer_visibility(map_document, "Risk_Zarar")
        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()
        # if stat[0] == u"q100":
        print "Exporting..", fname

        arcpy.mapping.ExportToPDF(map_document, fname)
        for elm in arcpy.mapping.ListLayoutElements(map_document, "TEXT_ELEMENT", "SubLegend_Address"):
            elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', u'TAŞKIN RİSK HARİTASI', elm.text, re.MULTILINE)


arcpy.env.overwriteOutput = True
start = datetime.datetime.now()
print start
mxdfiles = list([])
exemption = [
    "Ardahan_Risk.mxd",  # Done
    # "Akyaka_10_2__zarar.mxd",       #Done
    # "Akyaka_10_2__.mxd",            #Done
    # "Akyaka_10_2_2.mxd",            #Done
    # "Ardahan_10_2_zarar.mxd",       #Done
    # "Ardahan_10_2_.mxd",            #Done
    # "Arpacay_10_2_zarar.mxd",       #Done
    # "Arpacay_10_2_.mxd",            #Done
    # "CILDIR_10_2_zarar.mxd",        #Done
    # "CILDIR_10_2_.mxd",             #Done
    # "DOGUBAYAZIT_10_2_zarar.mxd",   #Done
    # "DOGUBAYAZIT_10_2_.mxd",        #Done
    # "GOLE_10_2_zarar.mxd",          #Done
    # "GOLE_10_2_.mxd",               #Done
    # "HANAK_10_2_zarar.mxd",         #Done
    # "HANAK_10_2_.mxd",              #Done
    # "HORASAN_10_2_zarar.mxd",       #Done
    # "HORASAN_10_2_.mxd",            #Done
    # "IGDIR_10_2_zarar.mxd",         #Done
    # "IGDIR_10_2.mxd",               #Done
    # "Karayazi_10_2_zarar.mxd",      #Done
    # "Karayazi_10_2_.mxd",           #Done
    # # "Kars_10_2_zarar.mxd",          #Done
    # "Kars_10_2_.mxd",               #Done
    # "KOPRUKOY_10_2_zarar.mxd",      #Done
    # "KOPRUKOY_10_2_.mxd",           #Done
    # "PASINLER_10_2_zarar.mxd",      #Done
    # "PASINLER_10_2_.mxd",           #Done
    # "SARIKAMIS_10_2_zarar.mxd",     #Done
    # "SARIKAMIS_10_2_.mxd",          #Done
    # "SELIM_10_2_zarar.mxd",         #Done
    # "SELIM_10_2_.mxd",              #Done
    # "TEKMAN_10_2_zarar.mxd",        #Done
    # "TEKMAN_10_2_.mxd",             #Done
]



for dirpath, subdirs, files in os.walk(process_folder):
    for x in files:
        if x.endswith("mxd") and x not in exemption:
            print x
            Q_intersect_layer = [["q50", u"Genel\TaskinSiniri\Q50 Taşkın Sınırı"],
                                 ["q100", u"Genel\TaskinSiniri\Q100 Taşkın Sınırı"],
                                 ["q500", u"Genel\TaskinSiniri\Q500 Taşkın Sınırı"]]

            aktivite_list = [u'Aktivite\Ticari Tesisler', u'Aktivite\Endüstriyel Tesisler',
                             u'Aktivite\Turistik Tesis Alanları', u'Aktivite\Trafolar',
                             u'Aktivite\Pazar Alanları', u'Aktivite\Lpg İstasyonları', u'Aktivite\Köprüler',
                             u'Aktivite\Yollar']

            stratejik_list = [u'Stratejik\IbadetHane', u'Stratejik\SaglikTesisleri', u'Stratejik\EgitimTesisleri']
            cevre_list = [u'Cevre\Ormanlık Alanlar', u'Cevre\Parklar', u'Cevre\Atiksu Arıtma Tesisleri']
            location_of_mxd = os.path.join(dirpath, x)
            path_of_process = os.path.dirname(location_of_mxd)
            excel_folder_name = os.path.abspath(os.path.join(path_of_process, "..", "EXCEL"))
            excel_file_name = glob.glob1(excel_folder_name, "*.xlsx")[0]
            output_path = os.path.abspath(os.path.join(path_of_process, "..", "PDF"))
            mxd = arcpy.mapping.MapDocument(location_of_mxd)
            temp_mxd = arcpy.mapping.MapDocument(location_of_mxd)
            # update_nufus_counts(mxd,'q50')
            # fieldname_check(mxd, u"BUILDING*", u"FIN_Q100")
            # mxd = arcpy.mapping.MapDocument("CURRENT")
            # temp_mxd = arcpy.mapping.MapDocument(location_of_mxd)
            excel_file = os.path.join(excel_folder_name, excel_file_name)
            workbook = xlrd.open_workbook(excel_file)
            worksheet = workbook.sheet_by_index(0)
            main_data_frame_name = "Veri"
            colormap_temp_layers = ["risk_template", "ekonomik_zarar_template"]
            df = arcpy.mapping.ListDataFrames(mxd, main_data_frame_name)[0]
            for lyr in arcpy.mapping.ListLayers(mxd):
                if lyr.name in colormap_temp_layers:
                    arcpy.mapping.RemoveLayer(df, lyr)

            risk_gr_layer = arcpy.mapping.Layer(settings_flag["Risk_Colormap"])
            zarar_gr_layer = arcpy.mapping.Layer(settings_flag["Ekonomik_Zarar_Colormap"])
            nufus_gr_layer = arcpy.mapping.Layer(settings_flag["nufus"])
            # arcpy.mapping.AddLayer(df, risk_gr_layer, "BOTTOM")
            # arcpy.mapping.AddLayer(df, zarar_gr_layer, "BOTTOM")
            arcpy.mapping.AddLayer(df, nufus_gr_layer, "TOP")


            for clr_mp_lyrs in colormap_temp_layers:
                layer_visibility(mxd, clr_mp_lyrs, False)
            remove_unwanted_layers_from_legends(mxd, colormap_temp_layers)
            Tobe_changed_risk_zarar = {"q_discharge_peak_table_title": u"KULLANILAN HİDROGRAF PİK DEĞERLERİ",
                                       "q_discharge_peak_address": u"Mustafakemalpaşa İlçe Merkezi",
                                       "q_discharge_peak_label": u"Q<SUB>100</SUB> (m<SUP>3</SUP>/sn )",
                                       "q_discharge_Peak_Value": u"1867.7"}

            layer_visibility(mxd, "Basemap", settings_flag["Basemap"])

            Input = {
                "q50": ["<SUB>50</SUB>", "NRSSN_Q50", "1489.8", u"1D\\Taşkın Alanları Su Yükseklikleri (m)(50)",
                        u"2D\\Taşkın Alanları Su Yükseklikleri (m)(50)", "FIN_Q50", "NRSSN_Q50"],
                "q100": ["<SUB>100</SUB>", "NRSSN_Q100", "1867.7", u"1D\\Taşkın Alanları Su Yükseklikleri (m)(100)",
                         u"2D\\Taşkın Alanları Su Yükseklikleri (m)(100)", "FIN_Q100", "NRSSN_Q100"],
                "q500": ["<SUB>500</SUB>", "NRSSN_Q500", "2712.5", u"1D\\Taşkın Alanları Su Yükseklikleri (m)(500)",
                         u"2D\\Taşkın Alanları Su Yükseklikleri (m)(500)", "FIN_Q500", "NRSSN_Q500"]
            }
            input_xl_data = []
            for sub_type in range(1, 4, 1):
                input_xl_data.append(get_input_values(worksheet, sub_type))

            # region Aktivite
            if settings_flag["Aktivite"]:

                smbl_lyr = arcpy.mapping.ListLayers(mxd, "BUILDING_RISK_v4")[0]
                arcpy.ApplySymbologyFromLayer_management(smbl_lyr, risk_gr_layer)
                layer_visibility(mxd, "Nufus", False)

                change_all_layers(mxd)
                change_antet_risk_zarar(mxd, Tobe_changed_risk_zarar, enabled=False)
                change_legend_gosterim(mxd, enabled=False)
                change_legend_zarar(mxd, enabled=False)
                change_legend_risk(mxd, mode="secondary")
                layer_visibility(mxd, "Risk_Zarar", visible=False)
                layer_visibility(mxd, "BUILDING_RISK_v4")
                change_data_frame_taskin_alani(mxd, enabled=False)

                arcpy.RefreshTOC()
                arcpy.RefreshActiveView()

                # Get Intersection Values of Each Segment


                change_legend_cevre_aktivite_stratejik(mxd, "Cevre", "Cevre", enabled=False)
                change_legend_cevre_aktivite_stratejik(mxd, "Aktivite", u"EKONOMİK AKTİVİTELER")
                layer_visibility(mxd, "Risk", visible=True)
                layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                for q in Q_intersect_layer:
                    val = get_intersection_values(mxd, aktivite_list, q[1])
                    run_risk(mxd, temp_mxd, [d for d in input_xl_data if d[0] == q[0]], mode='secondary')
                    for llx in arcpy.mapping.ListLayers(mxd):
                        if llx.longName.find(u"TaskinSiniri\\Q") != -1:
                            if llx.longName == q[1]:
                                # print llx.name, "will be enabled"
                                llx.visible = True
                            else:
                                # print llx.name, "will be disabled"
                                llx.visible = False
                    for lyr in arcpy.mapping.ListLayers(mxd, u"*Taşkın Sınırı İçinde Kalan Toplam Tesis Sayısı"):
                        temp = lyr.name
                        lyr.name = re.sub(re.compile('Q[0-9]{2,3}\sT', re.UNICODE), (q[0] + " T").upper(), temp)
                    for lyr in arcpy.mapping.ListLayers(mxd, u"*Adet"):
                        temp = lyr.name
                        lyr.name = re.sub(re.compile('^[0-9]{1,3}\sA', re.UNICODE), (str(val) + " A").upper(), temp)
                    print val
                    layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                    layer_visibility(mxd, "Risk", visible=True)
                    # change_oned_twod(mxd, q[0][1:])
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "PaftaNo"):
                        elm.text = [d for d in input_xl_data if d[0] == q[0]][0][12]
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubtypeDischarge"):
                        elm.text = re.sub('(?<=\/SUB\>\)\\r\\n)(.*?)(?=$)',
                                          "( " + [d for d in input_xl_data if d[0] == q[0]][0][18] + ")", elm.text,
                                          re.MULTILINE)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubLegend_Address"):
                        elm.text = re.sub('(?<=TASI)(.*?)(?=\(Q)',
                                          " / " + [d for d in input_xl_data if d[0] == q[0]][0][18], elm.text,
                                          re.MULTILINE)

                    arcpy.RefreshTOC()
                    arcpy.RefreshActiveView()
                    # arcpy.mapping.ExportToPDF(mxd, os.path.join(output_path, q[0] + "Aktivite" + ".pdf"))
                    arcpy.mapping.ExportToPDF(mxd, os.path.join(output_path,
                                                                x.split("_")[0] + "_Aktivite_" + q[0] + ".pdf"))
                    print "Exported succesfully"
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubLegend_Address"):
                        elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', u'TAŞKIN RİSK HARİTASI', elm.text, re.MULTILINE)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubtypeDischarge"):
                        elm.text = re.sub('(?<=\/SUB\>\)\\r\\n)(.*?)(?=$)', u" ", elm.text, re.MULTILINE)
            # endregion

            # region Cevre
            if settings_flag["Cevre"]:

                smbl_lyr = arcpy.mapping.ListLayers(mxd, "BUILDING_RISK_v4")[0]
                arcpy.ApplySymbologyFromLayer_management(smbl_lyr, risk_gr_layer)
                layer_visibility(mxd, "Nufus", False)

                change_legend_cevre_aktivite_stratejik(mxd, "Aktivite", "Aktivite", enabled=False)
                change_legend_cevre_aktivite_stratejik(mxd, "Cevre", u"ÇEVRESEL ZARARIN BOYUTU")
                layer_visibility(mxd, "Risk", visible=True)
                layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                for q in Q_intersect_layer:
                    val = get_intersection_values(mxd, cevre_list, q[1])
                    run_risk(mxd, temp_mxd, [d for d in input_xl_data if d[0] == q[0]], mode='secondary')
                    for llx in arcpy.mapping.ListLayers(mxd):
                        if llx.longName.find(u"TaskinSiniri\\Q") != -1:
                            if llx.longName == q[1]:
                                # print llx.name, "will be enabled"
                                llx.visible = True
                            else:
                                # print llx.name, "will be disabled"
                                llx.visible = False
                    for lyr in arcpy.mapping.ListLayers(mxd, u"*Taşkın Sınırı İçinde Kalan Toplam Tesis Sayısı"):
                        temp = lyr.name
                        lyr.name = re.sub(re.compile('Q[0-9]{2,3}\sT', re.UNICODE), (q[0] + " T").upper(), temp)
                    for lyr in arcpy.mapping.ListLayers(mxd, u"*Adet"):
                        temp = lyr.name
                        lyr.name = re.sub(re.compile('^[0-9]{1,3}\sA', re.UNICODE), (str(val) + " A").upper(), temp)
                    print val
                    layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                    layer_visibility(mxd, "Risk", visible=True)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "PaftaNo"):
                        elm.text = [d for d in input_xl_data if d[0] == q[0]][0][13]

                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubtypeDischarge"):
                        elm.text = re.sub('(?<=\/SUB\>\)\\r\\n)(.*?)(?=$)',
                                          "( " + [d for d in input_xl_data if d[0] == q[0]][0][19] + ")", elm.text,
                                          re.MULTILINE)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubLegend_Address"):
                        elm.text = re.sub('(?<=TASI)(.*?)(?=\(Q)',
                                          " / " + [d for d in input_xl_data if d[0] == q[0]][0][19], elm.text,
                                          re.MULTILINE)
                    arcpy.RefreshTOC()
                    arcpy.RefreshActiveView()
                    # arcpy.mapping.ExportToPDF(mxd, os.path.join(output_path, q[0] + "Cevre" + ".pdf"))
                    arcpy.mapping.ExportToPDF(mxd, os.path.join(output_path,
                                                                x.split("_")[0] + "_Cevre_" + q[0] + ".pdf"))
                    print "Exported succesfully"
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubLegend_Address"):
                        elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', u'TAŞKIN RİSK HARİTASI', elm.text, re.MULTILINE)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubtypeDischarge"):
                        elm.text = re.sub('(?<=\/SUB\>\)\\r\\n)(.*?)(?=$)', u" ", elm.text, re.MULTILINE)
            # endregion

            # region Stratejik
            if settings_flag["Stratejik"]:

                smbl_lyr = arcpy.mapping.ListLayers(mxd, "BUILDING_RISK_v4")[0]
                arcpy.ApplySymbologyFromLayer_management(smbl_lyr, risk_gr_layer)
                layer_visibility(mxd, "Nufus", False)
                remove_unwanted_layers_from_legends(mxd, ["mahalle_count", "mahalle_name"])

                change_legend_cevre_aktivite_stratejik(mxd, "Cevre", "Cevre", enabled=False)
                change_legend_cevre_aktivite_stratejik(mxd, "Stratejik", u"STRATEJİK TESİSLER")
                layer_visibility(mxd, "Risk", visible=True)
                layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)

                for q in Q_intersect_layer:
                    val = get_intersection_values(mxd, stratejik_list, q[1])
                    run_risk(mxd, temp_mxd, [d for d in input_xl_data if d[0] == q[0]], mode='secondary')
                    for llx in arcpy.mapping.ListLayers(mxd):
                        if llx.longName.find(u"TaskinSiniri\\Q") != -1:
                            if llx.longName == q[1]:
                                # print llx.name, "will be enabled"
                                llx.visible = True
                            else:
                                # print llx.name, "will be disabled"
                                llx.visible = False
                    for lyr in arcpy.mapping.ListLayers(mxd, u"*Taşkın Sınırı İçinde Kalan Toplam Tesis Sayısı"):
                        temp = lyr.name
                        lyr.name = re.sub(re.compile('Q[0-9]{2,3}\sT', re.UNICODE), (q[0] + " T").upper(), temp)
                    for lyr in arcpy.mapping.ListLayers(mxd, u"*Adet"):
                        temp = lyr.name
                        lyr.name = re.sub(re.compile('^[0-9]{1,3}\sA', re.UNICODE), (str(val) + " A").upper(), temp)
                    print val
                    layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                    layer_visibility(mxd, "Risk", visible=True)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "PaftaNo"):
                        elm.text = [d for d in input_xl_data if d[0] == q[0]][0][11]
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubtypeDischarge"):
                        elm.text = re.sub('(?<=\/SUB\>\)\\r\\n)(.*?)(?=$)',
                                          "( " + [d for d in input_xl_data if d[0] == q[0]][0][17] + ")", elm.text,
                                          re.MULTILINE)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubLegend_Address"):
                        elm.text = re.sub('(?<=TASI)(.*?)(?=\(Q)',
                                          " / " + [d for d in input_xl_data if d[0] == q[0]][0][17], elm.text,
                                          re.MULTILINE)
                    arcpy.RefreshTOC()
                    arcpy.RefreshActiveView()
                    layer_visibility(mxd, "Risk", visible=True)
                    # arcpy.mapping.ExportToPDF(mxd, os.path.join(output_path, q[0] + "Stratejik" + ".pdf"))
                    arcpy.mapping.ExportToPDF(mxd, os.path.join(output_path,
                                                                x.split("_")[0] + "_Stratejik_" + q[0] + ".pdf"))
                    print "Exported succesfully"
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubLegend_Address"):
                        elm.text = re.sub('(?<=^)(.*?)(?=\(Q)', u'TAŞKIN RİSK HARİTASI', elm.text, re.MULTILINE)
                    for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "SubtypeDischarge"):
                        elm.text = re.sub('(?<=\/SUB\>\)\\r\\n)(.*?)(?=$)', u" ", elm.text, re.MULTILINE)
            # endregion

            # region RISK
            if settings_flag["Risk"]:

                smbl_lyr = arcpy.mapping.ListLayers(mxd, "BUILDING_RISK_v4")[0]
                arcpy.ApplySymbologyFromLayer_management(smbl_lyr, risk_gr_layer)
                layer_visibility(mxd, "Nufus", False)
                remove_unwanted_layers_from_legends(mxd, ["mahalle_count", "mahalle_name"])
                layer_visibility(mxd, "ekonomik_zarar_template", visible=False)
                layer_visibility(mxd, "risk_template", visible=False)
                change_all_layers(mxd)
                change_legend_cevre_aktivite_stratejik(mxd, "Cevre", "Cevre", enabled=False)
                change_legend_gosterim(mxd)
                change_legend_zarar(mxd, enabled=False)
                change_legend_risk(mxd, mode='main')
                layer_visibility(mxd, "Risk")
                change_antet_risk_zarar(mxd, Tobe_changed_risk_zarar)
                change_data_frame_taskin_alani(mxd)
                # Run for Risk
                layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                layer_visibility(mxd, "Risk", visible=True)
                colormap_temp_layers = ["risk_template", "ekonomik_zarar_template"]
                df = arcpy.mapping.ListDataFrames(mxd, "Veri")[0]

                for lyr in arcpy.mapping.ListLayers(mxd):
                    if lyr.name in colormap_temp_layers:
                        arcpy.mapping.RemoveLayer(df, lyr)

                risk_gr_layer = arcpy.mapping.Layer( settings_flag["Risk_Colormap"])
                zarar_gr_layer = arcpy.mapping.Layer( settings_flag["Ekonomik_Zarar_Colormap"])
                arcpy.mapping.AddLayer(df, risk_gr_layer, "BOTTOM")
                arcpy.mapping.AddLayer(df, zarar_gr_layer, "BOTTOM")

                remove_unwanted_layers_from_legends(mxd, colormap_temp_layers)

                arcpy.RefreshTOC()
                arcpy.RefreshActiveView()
                run_risk(mxd, temp_mxd, input_xl_data)
                # run_risk({'q50': Input['q50']}, mode='secondary')
                # run_risk({'q500': Input['q500']})
            # endregion

            # region NUFUS
            if settings_flag["Nufus"]:

                smbl_lyr = arcpy.mapping.ListLayers(mxd, "BUILDING_RISK_v4")[0]
                arcpy.ApplySymbologyFromLayer_management(smbl_lyr, risk_gr_layer)
                # layer_visibility(mxd, "Nufus", False)
                layer_visibility(mxd, "Nufus", True)


                remove_unwanted_layers_from_legends(mxd, ["mahalle_count", "mahalle_name"])

                change_all_layers(mxd)
                change_legend_cevre_aktivite_stratejik(mxd, "Cevre", "Cevre", enabled=False)
                change_legend_gosterim(mxd)
                change_legend_zarar(mxd, enabled=False)
                change_legend_risk(mxd, mode='disable')
                layer_visibility(mxd, "Risk")
                change_antet_risk_zarar(mxd, Tobe_changed_risk_zarar)
                change_data_frame_taskin_alani(mxd)
                # Run for NUFUS
                layer_visibility(mxd, "BUILDING_RISK_v4", visible=False)
                layer_visibility(mxd, "Risk", visible=True)
                layer_visibility(mxd, "Risk_Zarar", visible=True)
                change_legend_nufus(mxd)
                run_nufus(mxd, input_xl_data)
                # run_risk({'q50': Input['q50']}, mode='secondary')
                # run_risk({'q500': Input['q500']})
            # endregion

            # region Zarar
            if settings_flag["Zarar"]:
                smbl_lyr = arcpy.mapping.ListLayers(mxd, "BUILDING_RISK_v4")[0]
                arcpy.ApplySymbologyFromLayer_management(smbl_lyr, zarar_gr_layer)
                layer_visibility(mxd, "Nufus", False)
                remove_unwanted_layers_from_legends(mxd, ["mahalle_count", "mahalle_name"])
                change_legend_nufus(mxd, enabled=False)

                change_all_layers(mxd)
                change_legend_cevre_aktivite_stratejik(mxd, "Cevre", "Cevre", enabled=False)
                change_legend_gosterim(mxd)
                change_legend_zarar(mxd)
                change_antet_risk_zarar(mxd, Tobe_changed_risk_zarar)
                change_legend_risk(mxd, enabled=False)
                layer_visibility(mxd, "Risk_Zarar")
                arcpy.RefreshActiveView()
                layer_visibility(mxd, "Risk")
                arcpy.RefreshTOC()
                change_data_frame_taskin_alani(mxd)
                # Run for Zarar
                layer_visibility(mxd, "BUILDING_RISK_v4", visible=True)
                layer_visibility(mxd, "Risk", visible=True)
                layer_visibility(mxd, "Raster", visible=True)

                # run_zarar(mxd, temp_mxd, Input)
                run_zarar(mxd, temp_mxd, input_xl_data)

            # endregion

print datetime.datetime.now() - start
