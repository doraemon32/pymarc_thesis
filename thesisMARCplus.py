from pymarc import MARCReader
from pymarc import Field
from configparser import ConfigParser
from datetime import datetime
import pandas as pd

# https://pypi.org/project/pymarc/
# https://carpentries-incubator.github.io/pymarc_basics/aio/index.html
# https://pymarc.readthedocs.io/en/latest/ =>textwrite
# https://groups.google.com/g/pymarc
# https://github.com/lpmagnuson/pymarc-workshop
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - thesisMARCplus.py get configure file
# ver 1.1
#  - bug fix
#    -write to text file using utf-8
#  - new requirement
#    -modify column to C041及新指令 & 546及原本546,content as follows
#       700 file 的 c546 有內容,判斷如下:
#       | c546 : "全中文"             |>=008 : chi >=041 $achi >無546
#       | c546 : "全英文"             |>=008 : eng >=041 $aeng >無546
#       | c546 : "中英對照"(或其他文字) |>=008 : eng >=041 $achi$aeng >=546 "中英對照"(或其他文字)
#       | c546 無內容: |>=008 : eng >=041 $achi$aeng >=546 "主要內容為英文"
#
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# define filename by datetime
now = datetime.now()
yymmddhh = now.strftime("%Y%m%d%H")
year = now.strftime('%G')
month = now.strftime('%m')
day = now.strftime('%d')
yymmdd = year[2:4] + month + day
# Read config.ini file
config_object = ConfigParser()
config_object.read("myconfig.ini")
# Get and display section
output_fileinfo = config_object["OUTPUTFILE"]
final_fileinfo = config_object["FINALOUTPUTFILE"]
double_fileinfo = config_object["DOUBLECHECK"]
# add datetime for outputfile name
# [OUTPUTPUTFILE] section
tempMRCfile = output_fileinfo["tempmrcfile"]
temp論文清單file = output_fileinfo["temp論文清單file"]
output0700file = output_fileinfo["0700mrcfile"]
tempMRCfilename_list = tempMRCfile.split("_")
# [FINALOUTPUTFILE] section
finalMRCfile = "final_" + tempMRCfilename_list[1] + "_fc_" + yymmddhh + ".mrc"
final_fileinfo["finalmrcfile"] = finalMRCfile
final論文清單file = "final_" + tempMRCfilename_list[1] + "_論文清單_" + yymmddhh + ".xlsx"
final_fileinfo["final論文清單file"] = final論文清單file
callNumberfile = "final_" + tempMRCfilename_list[1] + "_callnumber_" + yymmddhh + ".txt"
final_fileinfo["callnumberfile"] = callNumberfile
# [DOUBLECHECK] section
doubleCheckfile = "doubleCheck_" + tempMRCfilename_list[1] + "_" + yymmddhh + ".xlsx"
double_fileinfo["doubleCheckfile"] = doubleCheckfile
# Write changes back to ini file
with open("myconfig.ini", "w") as conf:
    config_object.write(conf)

# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - retrieve excel to dict list                                                       --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
columns_0700 = ['authorC','authorE','mentor1','mentor1E','mentor2','mentor2E','mentor3','mentor3E','mentor4','mentor4E','c546','原本546','author','titleNew','titleOrig']
columns_0700_Len = len(columns_0700)
pd0700_Df = pd.read_excel(output0700file, sheet_name=0, header=0, usecols=columns_0700)  # sheet 1st,one header
pd0700_Df = pd0700_Df.fillna("")    # fill all Nan value with ""
pd0700_dict_list = pd0700_Df.set_index('authorC').T.to_dict(orient='list')

columns_dissertation = ['系所', '索書號', '條碼', '書名', '作者', '出版項']
columns_dissertation_Len = len(columns_dissertation)
dissertation_Df = pd.read_excel(temp論文清單file, sheet_name=0, header=0,  usecols=columns_dissertation)  # sheet 1st,one header
dissertation_Df = dissertation_Df.fillna("")    # fill all Nan value with ""
pdDissertation_list = dissertation_Df.values.tolist()
pdDissertation_dict_list = dissertation_Df.set_index('作者').T.to_dict(orient='list')

# ----------------------------------------------procedure files -----------

# -------------Step: main update mrc data-----------------------
def processMarcPlus(record):
    # retrieve essential data
    author = record["100"]["a"]

    # ----------------------------------------------handle 0700.xlsx------------
    f546_new_dict.update({author: ""})
    f546_orig_dict.update({author: ""})
    title_orig_dict.update({author: ""})
    # Step : query author in the 0700 excel list and put in the list
    # 'authorE','mentor1','mentor1E','mentor2','mentor2E','mentor3','mentor3E','mentor4','mentor4E','c546','原本546','author','titleNew','titleOrig'
    if author in pd0700_dict_list:
        one_author_700s_list = pd0700_dict_list[author]

        record.remove_fields("700")
        # content get from excel
        content_n546 = str(one_author_700s_list[columns_0700.index("c546")-1])  # 'mentor4E'(8),'c546'(9),'原本546','author','titleNew'(12)
        content_new_title = str(one_author_700s_list[columns_0700.index("titleNew")-1])  # 'mentor4E'(8),'c546'(9),'原本546','author','titleNew'(12)
        # =0700 process
        for i in range(columns_0700.index("authorE")-1, columns_0700.index("mentor4E")):   # from 'authorE'(0) until 'mentor4E'(8)
            content_author = str(one_author_700s_list[i])
            if len(content_author) > 1:
                field_700 = Field(tag="700", indicators=["1", " "], subfields=["a", content_author])
                record.add_ordered_field(field_700)

        # =546 process
        if record["546"] is not None:
            field_546 = str(record["546"]["a"])
        else:
            field_546 = "無546"
        f546_orig_dict.update({author: field_546})

        record.remove_fields("008")
        record.remove_fields("041")
        record.remove_fields("546")
        if len(content_n546) > 1:
            # 700 file 的 c546 有內容,判斷如下:
            # | c546 : "全中文"             |>=008 : chi >=041 $achi >無546
            # | c546 : "全英文"             |>=008 : eng >=041 $aeng >無546
            # | c546 : "中英對照"(或其他文字) |>=008 : eng >=041 $achi$aeng >=546 "中英對照"(或其他文字)
            # | c546 無內容: |>=008 : eng >=041 $achi$aeng >=546 "主要內容為英文"
            if content_n546 == "全中文":
                # =008 process
                year = record["260"]["c"][0:4]  # remove last one char(.)
                ocn = yymmdd + "s" + year + "####ch#ad##e#b####000#0#" + "chi" + "#d"
                field_008 = Field(tag="008", data=ocn)
                # =041 process
                field_041 = Field(tag="041", indicators=[" ", " "], subfields=["a", "chi"])
                # =546 process
                #
            elif  content_n546 == "全英文":
                # =008 process
                year = record["260"]["c"][0:4]  # remove last one char(.)
                ocn = yymmdd + "s" + year + "####ch#ad##e#b####000#0#" + "eng" + "#d"
                field_008 = Field(tag="008", data=ocn)
                # =041 process
                field_041 = Field(tag="041", indicators=[" ", " "], subfields=["a", "eng"])
                # =546 process
                #
            else:
                # =008 process
                year = record["260"]["c"][0:4]  # remove last one char(.)
                ocn = yymmdd + "s" + year + "####ch#ad##e#b####000#0#" + "eng" + "#d"
                field_008 = Field(tag="008", data=ocn)
                # =041 process
                field_041 = Field(tag="041", indicators=["0", " "], subfields=["a", "chi", "a", "eng"])
                # =546 process
                field_546 = Field(tag="546", indicators=[" ", " "], subfields=["a", content_n546])
                record.add_ordered_field(field_546)
        else:
            content_n546 = ""
            # =008 process
            year = record["260"]["c"][0:4]  # remove last one char(.)
            ocn = yymmdd + "s" + year + "####ch#ad##e#b####000#0#" + "eng" + "#d"
            field_008 = Field(tag="008", data=ocn)
            # =041 process
            field_041 = Field(tag="041", indicators=["0", " "], subfields=["a", "chi", "a", "eng"])
            # =546 process
            field_546 = Field(tag="546", indicators=[" ", " "], subfields=["a", "主要內容為英文"])
            record.add_ordered_field(field_546)
            #
        record.add_ordered_field(field_008)
        record.add_ordered_field(field_041)
        #
        f546_new_dict.update({author: content_n546})

        # =245 process
        if len(content_new_title) > 1:
            orig_title = str(record["245"]["a"])   # part of title(245(a) and 245(b))
            if record["245"]["b"] is not None:
                orig_title = orig_title + " " + str(record["245"]["b"])   # complete title
            record.remove_fields("245")
            # 700 file 的 new_title 有內容,判斷如下:
            # new_title 有標準一個"="內容 | 應是 中文/英文 兩段  |>> =245 $a中文$b英文$c作者著
            # new_title 有標準一個"="內容 | 儰 中文/英文 兩段    |>> =245 $anew_title內容$c作者著 and 移除=246
            # new_title 沒有標準一個"="內容 |               |>> =245 $anew_title內容$c作者著 and 移除=246
            # new_title 超過標準一個"="內容 |Warning!!!     |>> =245 $anew_title內容$c作者著

            new_title_list = content_new_title.split("=")
            if len(new_title_list) == 2:
                if len(new_title_list[1]) > 2:
                    field_245 = Field(tag="245", indicators=["1", "0"],
                        subfields=["a", new_title_list[0]+"=", "b", new_title_list[1].lstrip(), "c", author+"著"])
                else:
                    #print("##2",new_title_list)
                    field_245 = Field(tag="245", indicators=["1", "0"],
                        subfields=["a", content_new_title, "c", author+"著"])
                    record.remove_fields("246")
            elif len(new_title_list) == 1:
                #print("##3",new_title_list)
                field_245 = Field(tag="245", indicators=["1", "0"], subfields=["a", content_new_title, "c", author+"著"])
                record.remove_fields("246")
            else:
                field_245 = Field(tag="245", indicators=["1", "0"], subfields=["a", content_new_title, "c", author+"著"])
                my_err = "**Warning::author "+author+" has strange new_title in the 0700 XLS file.**\n.==>>"+content_new_title
                err_list.append(my_err)
            record.add_ordered_field(field_245)

            title_orig_dict.update({author: orig_title})
    else:
        my_err = "**Error::mrc author "+author+" NotFound in the 0700 XLS file.**"
        err_list.append(my_err)

    # ------------------------------------------Handle temp論文清單.xlsx------------
    barcode_orig_dict.update({author: ""})
    callnumber_orig_dict.update({author: ""})
    # Step : query author from temp論文清單.xlsx,put in the list
    if author in pdDissertation_dict_list:
        one_author_dissertation_list = pdDissertation_dict_list[author]
        # '系所'(0),'索書號'(1),'條碼'(2),'書名'(3),'作者'(x),'出版項'(4)
        # content get from excel
        content_callnumber = str(one_author_dissertation_list[1])
        content_barcode = str(one_author_dissertation_list[2])

        # =084 process
        orig_callnumber = str(record["084"]["a"]) + " " + str(record["084"]["b"])
        if len(content_callnumber) > 1:
            if content_callnumber != orig_callnumber:
                try:
                    content_callnumber_list = content_callnumber.split(" ", 2)
                    record["084"]["a"] = content_callnumber_list[0]+" "+content_callnumber_list[1]
                    record["084"]["b"] = content_callnumber_list[2]
                    callnumber_orig_dict.update({author: orig_callnumber})
                except:
                    my_err = "**Warning::the callnumber content format is wrong!!!***"+author+content_callnumber
                    err_list.append(my_err)
                    content_callnumber_list = orig_callnumber.split(" ", 2)
                    record["084"]["a"] = content_callnumber_list[0]+" "+content_callnumber_list[1]
                    record["084"]["b"] = content_callnumber_list[2]
        else:
            my_err = "**Warning::the callnumber content is wrong!!***"+author
            err_list.append(my_err)

        # =035 process
        orig_barcode = str(record["035"]["a"])
        if orig_barcode != content_barcode:
            record["035"]["a"] = content_barcode
            barcode_orig_dict.update({author: orig_barcode})
    else:
        my_err = "**Error::mrc author "+author+" NotFound in the temp論文清單 XLS file.**"
        err_list.append(my_err)

    return record

# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - Main Program :                                                                    --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
err_list = []
my_marc_records = []
f546_new_dict = {}
f546_orig_dict = {}
title_orig_dict = {}
barcode_orig_dict = {}
callnumber_orig_dict = {}

# ------------------------------------------processMarcPlus process------------
with open(tempMRCfile, "rb") as fh:
    reader = MARCReader(fh)
    counter = 1

    for record in reader:
        print(counter, record["100"]["a"], record["245"]["a"])   # title
        singleRecord = processMarcPlus(record)
        counter = counter + 1
        my_marc_records.append(singleRecord)
        # ---Debug Information####
        '''
        for field in record:
            if "=095" in str(field):
                print("-->>",field)
        '''
# ----------------------------------------------output final mrcfile------------
with open(finalMRCfile, "wb") as out:
    for record in my_marc_records:
        # -- and write each record to it
        out.write(record.as_marc())

# --------------------------------------------------------------------------------------
# - output Program :                                                                  --
# --------------------------------------------------------------------------------------

# ----------------------------------------------procedure file ------------

# ----------------------------------------------output Items ------------
def outputItems(record):
    a_dissertation_list = []
    a_doublecheck_list = []

    grad = record["502"]["a"]
    author = record["100"]["a"]
    callnumber = record["084"]["a"] + " " + record["084"]["b"]
    edition = record["260"]["a"] + " " + record["260"]["b"] + record["260"]["c"]
    title = str(record["245"]["a"])
    if record["245"]["b"] is not None:
        title = title + " " + str(record["245"]["b"])
    else:
        my_err = "**Warning::mrc author "+author+" ~245 part $b~ not exists.**"
        err_list.append(my_err)
        #print(my_err)
    barcode = "abcdefg"
    if record["035"] is not None:
        barcode = str(record["035"]["a"])
    else:
        my_err = "**Warning::mrc author "+author+" barcode not exists.**"
        err_list.append(my_err)
        #print(my_err)

    a_dissertation_list.extend([grad, callnumber, barcode, title, author, edition])
    # above for 論文清單

    # more information ([[a_dissertation_list],name700_str,f_041,f_546,orig_callnumber,orig_barcode,orig_title])
    # orig f546,title,barcode,callNumber process
    orig_title = title_orig_dict[author]
    orig_barcode = barcode_orig_dict[author]
    orig_callnumber = callnumber_orig_dict[author]
    # name700_str process
    name700_str = author
    my_700_list = record.get_fields("700")
    for my_700 in my_700_list:
        name = my_700['a']  # e.g. =700  1\$aLee, Kung-Pei
        name700_str = name700_str + "/ " + name
    # f_041(C041及新指令) process
    f_041 = "無041"
    n0700_f546 = f546_new_dict[author]
    for field in record:
        if "=041" in str(field):
            f_041 = str(field).split(" ")[-1][2:]
    f_041 = f_041+"/"+n0700_f546

    # f_546(546及原本546) process
    f_546 = ""
    orig_f546 = f546_orig_dict[author]
    if record["546"] is not None:
        f_546 = record["546"]["a"]
    else:
        f_546 = "無546"
    f_546 = f_546+"/"+orig_f546
    #
    a_doublecheck_list.extend([grad,callnumber,barcode,title,author,edition,name700_str,f_041,f_546,orig_callnumber,orig_barcode,orig_title])
    # above for DoubleCheck

    return a_dissertation_list, a_doublecheck_list

# ------------------------------------------prepare output process------------
dissertation_list = []
doublecheck_list = []
prtCallnumber = ""

for record in my_marc_records:
    # 論文清單 doublecheck process
    a_dissertation_list, a_doublecheck_list = outputItems(record)
    dissertation_list.append(a_dissertation_list)
    doublecheck_list.append(a_doublecheck_list)
    # callNumber process
    grad_number_R = record["084"]["a"]
    fourCorner_list = record["084"]["b"].split(" ")
    # follow the spec format
    prtCallnumber = prtCallnumber+grad_number_R[:1]+"\n"+grad_number_R[-5:]+"\n"+fourCorner_list[0]+"\n"+fourCorner_list[1]+"\n\n\n\n\n\n\n\n"
#print(prtCallnumber)

# ----------------------------------------------output 3------------
# (1)
with open(callNumberfile, 'w', encoding='UTF-8') as f:  # callNumber output process
    f.write(prtCallnumber)
# (2)
columns_df2 = ['系所','索書號','條碼','書名','作者','出版項','指導者','C041及新指令','546及原本546','原本索書號','原本條碼','原本書名']
df2 = pd.DataFrame(doublecheck_list, columns=columns_df2)
df2.index.name = "項次"
df2.index += 1
df2.to_excel(doubleCheckfile)
# (3)
# columns_dissertation = ['系所', '索書號', '條碼', '書名', '作者', '出版項']
df1 = pd.DataFrame(dissertation_list, columns=columns_dissertation)
df1.index.name = "項次"
df1.index += 1
df1.to_excel(final論文清單file)
# --------------------------------------------------------------------------------------
if err_list:
    print("\n")
    for e in err_list:
        print(e)
# --------------------------------------------------------------------------------------

