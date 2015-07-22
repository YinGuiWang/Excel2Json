#!/usr/bin/python
__author__ = 'yogi'

import json
import xlrd
import os,math,types
from collections import OrderedDict,namedtuple
import sys

global gError_log

def PrinFuncAndLine(*args):
    try:
       raise Exception
    except:
         f = sys.exc_info()[2].tb_frame.f_back
    #print  '%s, %s, %s, :' % (f.f_code.co_filename, f.f_code.co_name, str(f.f_lineno)),args

    gError_log.write(args.__str__() + '\n')


def ReviseValue(value):
    result = value
    try:
        result = float(value);
        if math.floor(result) == result:
            result = int(result)
    except Exception,e:
        ""
        #PrinFuncAndLine(e)
    finally:
        return result


global g_KeyCol, g_FirstIdx
g_KeyCol = 0
g_FirstIdx = 2

class ExcelToJson:


    RowIdx4Key = {}

    def __init__(self, cfgJsonPath):

        if not os.path.isfile(cfgJsonPath):
            return

        congfigFp =  file(cfgJsonPath)
        cfgList = json.load(congfigFp, object_pairs_hook = OrderedDict)
        self.fileCfgTb = cfgList[0]
        self.ContentCfgTb = cfgList[1]
        #print self.ContentCfgTb
        self.StartDoExcel2Json()


    def StartDoExcel2Json(self):

        excel_path =  self.fileCfgTb["excel_path"]
        excel_path = excel_path.replace("/", "\\")
        main_sheet = self.fileCfgTb["main_sheet"]

        allResult = []
        try:
            self.Excle_Object = xlrd.open_workbook(excel_path)
            sheet_info = self.Excle_Object.sheet_by_name(main_sheet)

            for rowNum in range(g_FirstIdx, sheet_info.nrows):
               primaryKey = sheet_info.cell(rowNum,0).value
               #print "%d"% (primaryKey)
               result = self.PackRowxData(primaryKey)
               allResult.append(result)
        except Exception,e:
            PrinFuncAndLine(e)
        finally:
            if self.fileCfgTb["is_merge"]:
                try:
                    outpuy_path = self.fileCfgTb["output_path"]
                    outpuy_path = outpuy_path.replace("/", "\\")
                    if not os.path.exists(outpuy_path):
                        os.makedirs(outpuy_path)
                    outfile = open(outpuy_path + "\\output.json", "w" )
                    outfile.write(json.dumps(allResult, indent=1))
                    outfile.close()
                    #print json.dumps([resultObject], indent=1)
                except Exception,e:
                    PrinFuncAndLine(e)
            print "Success!"


    #pack row data
    def PackRowxData(self,primaryKey):

        resultObject = OrderedDict()
        try:
            for tbKey, tbValue in self.ContentCfgTb.items():
               #print tbValue, type(tbValue)
               self.AnalyticalDispatch(resultObject, tbKey, tbValue, primaryKey)
        except Exception,e:
            PrinFuncAndLine(e)

        finally:
            try:
                outpuy_path = self.fileCfgTb["output_path"]
                outpuy_path = outpuy_path.replace("/", "\\")
                if not os.path.exists(outpuy_path):
                    os.makedirs(outpuy_path)
                outfile = open(outpuy_path + "\\" + str(ReviseValue(primaryKey)) + ".json", "w" )
                outfile.write(json.dumps([resultObject], indent=1))
                outfile.close()
                #print json.dumps([resultObject], indent=1)
            except Exception,e:
                PrinFuncAndLine(e)

        return resultObject


    #
    def AnalyticalDispatch(self,resultObject, tbKey, tbValue, primaryKey):

        #if type(tbValue) is types.DictionaryType:
        if isinstance(tbValue, OrderedDict):
            self.AnalyticalDic(resultObject, tbKey, tbValue, primaryKey)
        elif type(tbValue) is types.ListType:
            self.AnalyticaList(resultObject, tbKey, tbValue, primaryKey)


    #List
    def AnalyticaList(self,resultObject, tbKey, tbValue, primaryKey):
        if not type(tbValue) is types.ListType:
            return
        resultObject[tbKey] = []

        list_info = tbValue[0]

        if list_info["type"] == 1:
            rowList = self.GetRowIdxBySheetWithKey( list_info["sheet"],  list_info["commonCol"], primaryKey)
            #self.AnalyticalDispatch(rowIdx, value)
            sheet_info = self.Excle_Object.sheet_by_name(list_info["sheet"])
            for rowIdx in rowList:
                #print "rowIdx = %d" %(rowIdx)
                newObject = OrderedDict()
                newPrimaryKey = sheet_info.cell(rowIdx, g_KeyCol).value
                self.AnalyticalDic(newObject, tbKey, tbValue[1], newPrimaryKey)
                resultObject[tbKey].append(newObject)

        elif list_info["type"] == 2:

            sheet_info = self.Excle_Object.sheet_by_name(list_info["sheet"])
            rowList = self.GetRowIdxBySheetWithKey( list_info["sheet"], g_KeyCol, primaryKey)
            valueStr = sheet_info.cell(rowList[0],list_info["commonCol"]).value
            valueList = valueStr.split(";")

            #resultList = []
            for value in valueList:
                value = value.strip('()')
                #print  value
                if value == '':
                    continue
                fValueList = value.split(',')
                if len(fValueList) != len(list_info["keys"]):
                    continue
                #print  fValueList
                result = OrderedDict()
                i = 0
                for fValue in fValueList:
                    try:
                        result[list_info["keys"][i]] = float(fValue)
                    except:
                        result[list_info["keys"][i]] = fValue
                    finally:
                        i = i + 1
                resultObject[tbKey].append(result)
            #print PrinFuncAndLine(resultList)


    #Dic
    def AnalyticalDic(self, resultObject, tbKey, tbValue, primaryKey):

        if not isinstance(tbValue, OrderedDict):
            return
        #PrinFuncAndLine(primaryKey)
        if  tbValue.has_key("sheet")  and tbValue.has_key("valueCol") and tbValue.has_key("result"):  # and tbValue.has_key("keyCol")
            #print tbValue
            self.AnalyticalFinalDic(resultObject, tbKey, tbValue, primaryKey)
        else:
            for newTbKey, newTbValue in tbValue.items():
                self.AnalyticalDispatch(resultObject, newTbKey, newTbValue, primaryKey)


    #final data
    def AnalyticalFinalDic(self, resultObject, tbKey, json_data, valueKey):

        sheet_name = json_data["sheet"]
        sheet_info = self.Excle_Object.sheet_by_name(sheet_name)
        rowList = self.GetRowIdxBySheetWithKey( json_data["sheet"],  g_KeyCol, valueKey)

        if len(rowList) == 1:
            value = sheet_info.cell(rowList[0],json_data["valueCol"]).value
            if value == '' or value == None:
                value = 0
            other = []
            i = 0
            while(json_data.has_key("other[" + str(i) + "]")):
                other.append(sheet_info.cell(rowList[0],json_data["other[" + str(i) + "]"]).value)
                #print "value = " + str(value)
                #print "curOther = ",other[i]
                i = i + 1
            resultObject[tbKey] = ReviseValue(eval(json_data["result"]))
        else:
            print "Json Data Error! Here:"
            print json_data


    #get row index
    def GetRowIdxBySheetWithKey(self, sheet, keyCol, valueKey):

        if not self.RowIdx4Key.has_key(sheet):
            self.RowIdx4Key[sheet] = {}
            #print  self.RowIdx4Key
        if not self.RowIdx4Key[sheet].has_key(keyCol):
            self.RowIdx4Key[sheet][keyCol] = {}
            sheet_info = self.Excle_Object.sheet_by_name(sheet)
            for row in range(g_FirstIdx,sheet_info.nrows):
                primaryKey = sheet_info.cell(row,keyCol).value
                if not self.RowIdx4Key[sheet][keyCol].has_key(primaryKey):
                    self.RowIdx4Key[sheet][keyCol][primaryKey] = []
                self.RowIdx4Key[sheet][keyCol][primaryKey].append(row)

        return self.RowIdx4Key[sheet][keyCol][valueKey]



if __name__ == '__main__':

    dest = raw_input("Please Input Config Path:")
    if dest == "":
        dest = "Config.json"
    
    gError_log = open("error_log.txt", "w" )
    ExcelToJson(dest)
    gError_log.close()