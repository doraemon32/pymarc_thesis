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
# - thesisMARC.py 1st phase get configure file                                        --
# ver 1.1
#  - bug fix
#    -fourcorner serial is equal to '10','20'.. case
#    -write to text file using utf-8
#  - mark : to be added here for finding missing or duplicate fourcorner serial
#  - new requirement
#    -add new column '原本546' and the related content in 0700 file
#    -orig marc data not found in processMarc, e.g.041$a not exist,print all contents
# ver 1.2
#  - new requirement
#    -no need for 095 now so 'remove field_095'
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# define filename by datetime
now = datetime.now()
yymmddhh = now.strftime("%Y%m%d%H")
# Read config.ini file
config_object = ConfigParser()
config_object.read("myconfig.ini")
# Get and display section
myinfo = config_object["DEFAULT"]
input_fileinfo = config_object["INPUTFILE"]
output_fileinfo = config_object["OUTPUTFILE"]
# [DEFAULT] section
myName = myinfo["mynamefor095"]
myProcessTime = myinfo["myprocesstimefor035"]
# [INPUTFILE] section
origMRCfilename = input_fileinfo["origmrcfile"]
barCodefilename = input_fileinfo["barcodefile"]
fourCornerSerialfilename = input_fileinfo["fourcornerserialfile"]
# add datetime for outputfile name
origMRCfilename_list = origMRCfilename.split(".")
# [OUTPUTPUTFILE] section
tempMRCfile = "tmp_" + origMRCfilename_list[0] + "_" + yymmddhh + ".mrc"
output_fileinfo["tempmrcfile"] = tempMRCfile
temp論文清單file = "tmp_" + origMRCfilename_list[0] + "論文清單_" + yymmddhh + ".xlsx"
output_fileinfo["temp論文清單file"] = temp論文清單file
output0700file = "0700_" + origMRCfilename_list[0] + "_" + yymmddhh + ".xlsx"
output_fileinfo["0700mrcfile"] = output0700file
missMRCfile = "miss_" + origMRCfilename_list[0] + "_" + yymmddhh + ".mrc"
output_fileinfo["missmrcfile"] = missMRCfile
errorMRCfile = "error_" + origMRCfilename_list[0] + "_" + yymmddhh + ".log"
output_fileinfo["errormrcfile"] = errorMRCfile
# Write changes back to ini file
with open("myconfig.ini", "w") as conf:
    config_object.write(conf)

# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - retrieve barCodefilename excel to dict list                                       --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# get XLSX file.#item,Author,Title,Barcode,four_corner,Appendix
barCode_Df = pd.read_excel(barCodefilename, sheet_name=0, header=0, usecols=[0, 1, 2, 3, 4, 5, 6, 7])  # sheet 1st,one header
barCode_Df.columns = ["item","name","title","four_corner","barcode","appendix","c546","new_title"]  # item,Author,Title,Barcode,four_corner,Appendix,c546,new_title
barCode_Df = barCode_Df.fillna("")    # fill all Nan value with ""
len_of_barcode_xlsx = len(barCode_Df.index)
name_barcode_dict_list = barCode_Df.set_index('name').T.to_dict(orient='dict')
item_name_dict_list = barCode_Df.set_index('item').T.to_dict(orient='dict')
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - retrieve barCodefilename excel to dict list                                       --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# get XLSX file.#item,Author,Title,Barcode,four_corner,Appendix
fourCorner_Df = pd.read_excel(fourCornerSerialfilename, sheet_name=0, header=0)
fourCorner_list = fourCorner_Df[fourCorner_Df.columns[0]].tolist()

def lst2DictConvert(lst_orig):
    res_dct = {}    # result of dict : {str-int : max of decimal} for four corner map to series number
    decimal_list = []
    guessMax = lst_orig[0].split('.')[0]  # judge for begin
    guessMax_decimal = 0
    lst_orig.append(float(99999.9))  # judge for ending

    for item in lst_orig:

        # Step : Split Integer and Decimal
        int_decimal_list = str(item).split('.')
        int_part = int_decimal_list[0]
        if len(int_decimal_list) == 1:
            decimal_part = 0
        else:
            decimal_part = int(int_decimal_list[1])

        if guessMax < int_part:
            # Step : put max of Decimal part into the dict
            res_dct.update({guessMax : guessMax_decimal})
            # to be added here for finding missing or duplicate fourcorner serial #
            # Step : condition init
            guessMax = int_part
            guessMax_decimal = decimal_part
            decimal_list = []
        else:
            if guessMax_decimal < decimal_part:
                guessMax_decimal = decimal_part
        decimal_list.append(decimal_part)

    return res_dct

def getFourCornerXlsx():    # access permanent callNumber download file

    list0088 = []
    list0089 = []
    callnumber0088_list = []
    callnumber0089_list = []
    for item in fourCorner_list:
        if item.split(' ')[1] == "008.8":   # based on 2 grad_number,put full four Corner to these 2 lists
            list0088.append(str(item.split(' ')[2]))
        else:
            list0089.append(str(item.split(' ')[2]))
    # Step : finding max by sorting method
    list0088_sorted = sorted(list0088, key=float)  # in case if the values were there as string.
    list0089_sorted = sorted(list0089, key=float)  # in case if the values were there as string.

    if list0088_sorted:
        callnumber0088_list = lst2DictConvert(list0088_sorted)  # list2dict and remove duplicate key(integer)
    if list0089_sorted:
        callnumber0089_list = lst2DictConvert(list0089_sorted)

    return callnumber0088_list, callnumber0089_list

# define fixed strings
s008 = "=008  220103s2021\\\\ch\ad\\e\b\\\\000\0\chi\d"  # 220103 (now) #chi or eng (get from =041)
s020 = "=020  \\\$q(精裝)"    # new item
s035 = "=035  \\$a30010101877952$b2021/11/01$kR$h008.8"
# barcode #碩士(008.8)(碩士 from =502) #$b2021/11/01 fix date
s040 = "=040  \\$aNYCU$bchi$cNYCU$eccr"  # YMU->NYCU
s041 = "=041  \\$achi"  # =041  \\$aeng => =041  0\$achi$aeng # =041  \\$achi => same
s044 = "=044  \\$ach"   # new item
s084 = "=084  \\$aR008.8$b6624 2015$2ncsclt"  # 碩士(008.8)(碩士 from =502) #four_corner(6624) #year(2015 from=260)
s095 = "=095  \\$aYMU$s220103$wkuanwu8$9local"  # new item #220103 (now)
s260 = "=260  \\$a臺北市 :$b嚴緒芳,$c2015"  # 嚴緒芳 (from =100)
s300 = "=300  \\$a66頁 :$b圖, 表$c30公分"  # add one space
s500 = "=500  \\$a含附錄"  # \\$a含附錄 or '附錄: 1, 內容 ; 2, 內容' or 'null' #????
s502 = "=502  \\$a碩士--國立陽明大學生化暨分子生物研究所, 2015"   # " 02015" (year from =260)
s546 = "=546  \\$a主要內容為英文"  # new item if # =041  \\$aeng
s592 = "=592  \\"   # remove item
s700a = "=700  1\$aYen, Hsu-Fang"
s700b = "=700  1\$a蔡英傑"
s700c = "=700  1\$aTsai, Ying-Chieh"
s902 = "=902  \\"   # remove item
s994 = "=994  \\"   # remove item
s999 = "=999  \\"   # remove item
year = now.strftime('%G')
month = now.strftime('%m')
day = now.strftime('%d')
yymmdd = year[2:4] + month + day

fourCorner0088_dict, fourCorner0089_dict = getFourCornerXlsx()

# -------------Step1: procedures-----------------------

def get_four_corner_series(query, grad_number):  # eg. query='0041'(four corner number),grad_number = '008.8'
    # e.g. return 8(means 0041.0~0041.7 exists)
    four_corner_series = 9999   # in case not exist case, return value

    if grad_number == "008.8":
        if fourCorner0088_dict:
            if query in fourCorner0088_dict:  # e.g.0041 exists
                four_corner_series = fourCorner0088_dict[query]+1
            else:  # e.g.0041 not exists
                four_corner_series = 0
            # Step : update or add to dict
            newvalue_dic = {query: four_corner_series}
            fourCorner0088_dict.update(newvalue_dic)
    else:
        if fourCorner0089_dict:
            if query in fourCorner0089_dict:  # e.g.0041 exists
                four_corner_series = fourCorner0089_dict[query]+1
            else:  # e.g.0041 not exists
                four_corner_series = 0
            # Step : update or add to dict
            newvalue_dic = {query: four_corner_series}
            fourCorner0089_dict.update(newvalue_dic)

    return four_corner_series

def getBarcodeAndSoOn(query_author, title_mrc, grad_number):
    # ['item','name','title','four_corner','barcode','appendix','c546','new_title']
    # init values for return value in case empty value
    item_number = 999
    barcode_number = "99999999999999"
    appendix_number = "x"
    four_corner_number = "9999"
    four_corner_series = 999

    if query_author in name_barcode_dict_list:    # get['項次','書名','條碼','索書號'] based on '作者'
        item_number = name_barcode_dict_list[query_author]['item']
        barcode_number = str(name_barcode_dict_list[query_author]['barcode'])
        appendix_number = name_barcode_dict_list[query_author]['appendix']  # 含附錄
        if len(appendix_number) < 1:
            appendix_number = "x"
        four_corner_number = str(name_barcode_dict_list[query_author]['four_corner']).zfill(4)
        four_corner_series = get_four_corner_series(four_corner_number, grad_number)

        title_number = name_barcode_dict_list[query_author]['title']
        if not (title_mrc in title_number):  # '吳孟修' exist,but '書名' different,keep output,print error for warning**
            my_err = "**Warning::different title,name:"+query_author+"(item:"+str(item_number)+")\n\torig mrc:"+title_mrc+"\n\tnew title:"+title_number
            err_list.append(my_err)
            #print(my_err)
    else:
        my_err = "**Error::mrc author "+query_author+" NotFound in the XLS file.**"
        err_list.append(my_err)
        #print(my_err)

    # return essential items
    return item_number, barcode_number, appendix_number, four_corner_number, four_corner_series

# -------------Step: main update mrc data-----------------------

def processMarc(record):
    # retrieve essential data
    try:
        lang = record["041"]["a"]
        year = record["260"]["c"][0:4]  # remove last one char(.)
        grad = record["502"]["a"][0:2]  # e.g. "碩士"
        author = record["100"]["a"]
        title = record["245"]["a"]  # part of title(245(a) and 245(b))
    except:
        my_err = "**Error::some field is none,please check : [041][a]/[260][c]/[502][a]/[100][a]/[245][a].**"
        err_list.append(my_err)
        print(record)
        return record, 999

    if grad == "碩士":
        grad_number = "008.8"
        grad_number_R = "R 008.8"
    else:
        grad_number = "008.9"
        grad_number_R = "R 008.9"

    itemOrder_inXls, barcode, appendix, fourCorner, fourCornerSeries = getBarcodeAndSoOn(author, title, grad_number)

    if itemOrder_inXls == 999:  # itemOrder_inXls=999 means the mrc-name not found in namebarcode1.xls
        return record, itemOrder_inXls

    # ----------------CRUD mrc record now---------------------

    record.remove_fields("008")
    record.remove_fields("020")
    record.remove_fields("035")
    record.remove_fields("040")
    record.remove_fields("041")
    record.remove_fields("044")
    record.remove_fields("084")
    record.remove_fields("095")
    record.remove_fields("546")
    record.remove_fields("592")
    record.remove_fields("902")
    record.remove_fields("994")
    record.remove_fields("999")

    ocn = yymmdd + "s" + year + "####ch#ad##e#b####000#0#" + lang + "#d"
    field_008 = Field(tag="008", data=ocn)
    # =020 $q(精裝)
    field_020 = Field(tag="020", indicators=[" ", " "], subfields=["q", "(精裝)"])
    # =035  \\$a30010101877952$kR$h008.8$i2611.12 2015
    if fourCornerSeries != 0:
        four_corner = str(fourCorner) + "." + str(fourCornerSeries) + " " + year
    else:
        four_corner = str(fourCorner) + " " + year
    field_035 = Field(
        tag="035",
        indicators=[" ", " "],
        subfields=[
            "a", barcode,
            "b", myProcessTime,
            "k", "R",
            "h", grad_number,
        ],
    )
    # =040 $aNYCU$bchi$cNYCU$eccr
    field_040 = Field(
        tag="040",
        indicators=[" ", " "],
        subfields=[
            "a", "NYCU",
            "b", "chi",
            "c", "NYCU",
            "e", "ccr",
        ],
    )

    # =041  \\$aeng => =041  0\$achi$aeng
    # =041  \\$achi => same
    if lang == "eng":
        field_041 = Field(
            tag="041", indicators=["0", " "], subfields=["a", "chi", "a", "eng"]
        )
    else:
        field_041 = Field(
            tag="041", indicators=[" ", " "], subfields=["a", lang]
        )
    # =044 \\$ach
    field_044 = Field(tag="044", indicators=[" ", " "], subfields=["a", "ch"])
    # =084 \\$aR 008.8$bk$2ncsclt
    field_084 = Field(
        tag="084",
        indicators=[" ", " "],
        subfields=["a", grad_number_R, "b", four_corner, "2", "ncsclt"],
        # =502  \\$a碩士--國立陽明大學生醫光電研究所, 2020 => R 008.8
        # =502  \\$a博士--國立陽明大學生醫光電研究所, 2020 => R 008.9
    )
    # =095  \\$aYMU$s220107$wkuanwu8$9local
    '''field_095 = Field(
        tag="095",
        indicators=[" ", " "],
        subfields=["a", "YMU", "s", yymmdd, "w", myName, "9", "local"],
    )'''
    # =546
    field_546 = Field(tag="546", indicators=[" ", " "], subfields=["a", "主要內容為英文"])
    if lang == "eng":   # only eng case needs field_546
        record.add_ordered_field(field_546)

    record.add_ordered_field(
        field_008,
        field_020,
        field_035,
        field_040,
        field_041,
        field_044,
        field_084,
        field_095,
    )
    # update value
    record["260"]["b"] = author + ", "
    record["300"]["b"] = "圖, 表"
    record["502"]["a"] = record["502"]["a"] + ", " + year

    # =500 process
    a_list = []
    my_500s = record.get_fields('500')
    for my_500 in my_500s:
        my_500_value = my_500['a']
        a_list.append(my_500_value)
    # \\$a含附錄 always in the (=500) last item????? how to check otherwise del essential 500??
    if "附錄" or "刪除" in a_list[-1]:
        del a_list[-1]  # last one item "=500  \\$a含附錄"
    else:   # e.g. 3. Warning : last one of =500 is : 紙本論文延後公開(編目時請刪除本段)
        my_err = "**Error::last one item =500: "+a_list[-1]+".by author : "+author+"**"
        err_list.append(my_err)
        #print(my_err)
    # "x" means : no more appendix information
    if not appendix.lower() == "x":
        my_500_newvalue = appendix
        a_list.append(my_500_newvalue)

    record.remove_fields("500")
    for item_appendix in a_list:
        field_500 = Field(tag="500", indicators=[" ", " "], subfields=["a", item_appendix])
        record.add_ordered_field(field_500)

    # =700 process (handle eng/chi naming order)
    a_list = []
    b_list = []
    c_list = []
    my_700s = record.get_fields('700')
    for my_700 in my_700s:
        name = my_700['a'].title()  # Capitalize the first letter of every word in the list
        # reverse direction partiton function(split two partions) to list
        name1 = name.rpartition(' ')
        # get one char for later analysis
        _char = name1[-1][0:1]
        if not '\u4e00' <= _char <= '\u9fa5':   # is eng
            LastName = name1[-1].split('.')[0] + ", " + name1[0]
            #print(name,'is Eng')
        else:
            LastName = name
        a_list.append(LastName)

    i_eng = 0
    i_chi = 1
    # based on a_list,create b_list for the order of a_list
    # e.g. a_list=[c0,c1,a,e0,e1], result c_list=[a,c0,e0,c1,e1]
    #   create b_list=[1,3,0,2,4] order
    # i.e a_list=[c0,c1,a,e0,e1]/b_list=[1,3,0,2,4] => c_list=[a,c0,e0,c1,e1]
    for item in a_list:
        _char = item[0:1]
        if not '\u4e00' <= _char <= '\u9fa5':   # is eng
            b_list.append(i_eng)
            i_eng = i_eng + 2
        else:
            b_list.append(i_chi)
            i_chi = i_chi + 2
    #print(author,'\n',a_list,'\n',b_list)
    for i in range(i_eng-1):
        #print("------->>>>",i,b_list.index(i),a_list[b_list.index(i)])
        if i in b_list:
            item = a_list[b_list.index(i)]
            c_list.append(item)

    record.remove_fields("700")
    for name in c_list:
        field_700 = Field(tag="700", indicators=["1", " "], subfields=["a", name])
        record.add_ordered_field(field_700)

    return record, itemOrder_inXls  # itemOrder_inXls for later in-sequence purpose

# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - main program :                                                                    --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
err_list = []
my_marc_records_updated = []
my_marc_records_miss_barcode = []
my_marc_sequence = []

# -------------Step1: original MRC file check and update-----------------------
with open(origMRCfilename, "rb") as fh:
    reader = MARCReader(fh)
    counter = 1
    counter_miss = 0

    for record in reader:
        print(counter, record['245']['a'])   # title
        counter = counter + 1
        #print(record)
        # itemOrder_inXls for later in-sequence purpose
        singleRecord, itemOrder_inXls = processMarc(record)
        if itemOrder_inXls == 999:
            counter_miss = counter_miss + 1
            my_marc_records_miss_barcode.append(singleRecord)
        else:
            my_marc_records_updated.append(singleRecord)
            my_marc_sequence.append(itemOrder_inXls)
        #print(record)
        #print("====================================")

# -------------Step1.1: barcode not updated part-----------------------
if my_marc_records_miss_barcode:
    with open(missMRCfile, "wb") as out:
        for record in my_marc_records_miss_barcode:
            # -- and write each record to it                 #
            out.write(record.as_marc())
# -------------missing part end ------------------------------

# -------------Step1.2: barcode already updated part-----------------------
# len_of_barcode_xlsx = len(barCode_Df.index)
my_marc_records = []    # ready to store the arrange sequence mrac records(from my_marc_records_updated)
for index_of_xls in range(1, len_of_barcode_xlsx+1):

    authorFromXls = item_name_dict_list[index_of_xls]['name']
    if index_of_xls not in my_marc_sequence:
        # select_information = namebarcode_Df.loc[namebarcode_Df['item'] == i]
        my_err = "**Error Index "+str(index_of_xls)+":"+authorFromXls+" does not exist in the mrc file.**"
        err_list.append(my_err)
        #print('\nList the name sequence number from~',barCodefilename,'file,dependent on~',origMRCfilename,'file sequence.\n',my_marc_sequence)
        #print(my_err)
    else:
        tmp_record = my_marc_records_updated[my_marc_sequence.index(index_of_xls)]
        my_marc_records.append(tmp_record)

with open(tempMRCfile, "wb") as out:
    for record in my_marc_records:
        # -- and write each record to it
        out.write(record.as_marc())

# -------------barcode already updated part end--------------------

# -------------Step1.3: Error log into file-----------------------
with open(errorMRCfile, 'a', encoding='UTF-8') as f:
    prtError = "-----------------------" + str(now) + "----------------------\n"
    prtWarn = ""
    for item in err_list:
        if "**Err" in item:
            prtError = prtError + item + "\n"
        else:
            prtWarn = prtWarn + item + "\n"
    prtError = prtError + "\n\n" + prtWarn
    f.write(prtError)
# -------------original MRC file check and update end-----------------------

# -------------Step2: create 0700&論文清單 for check by manually-----------------------
columns_0700 = ['authorC','authorE','mentor1','mentor1E','mentor2','mentor2E','mentor3','mentor3E','mentor4','mentor4E','c546','原本546','author','titleNew','titleOrig']
columns_0700_Len = len(columns_0700)
columns_論文清單 = ['系所', '索書號', '條碼', '書名', '作者', '出版項']
dissertation_len = len(columns_論文清單)

# -------------Step2.1: procedures-----------------------
def getC546NewTitle(query_author, title_mrc):
    # ['item','name','title','four_corner','barcode','appendix','c546','new_title']
    # init values for return value in case empty value
    # query_author = '吳孟修' #title_mrc = '利用光致電流影像(OBIC)研究'
    c546_number = ""
    title_new_number = ""

    if query_author in name_barcode_dict_list:    # get['title','c546','new_title'] based on 'name'
        c546_number = name_barcode_dict_list[query_author]['c546']
        title_number = name_barcode_dict_list[query_author]['title']
        title_new_number = name_barcode_dict_list[query_author]['new_title']

        if not (title_mrc in title_number):  # '吳孟修' exist,but '書名' different,keep output,print error for warning**
            my_err = "**Warning::different title,name:"+query_author+".\n\t orig mrc:"+title_mrc+"\n\tnew title:"+title_number
            err_list.append(my_err)
            print(my_err)

    return c546_number, title_new_number

def outputItems(my_record):
    # 論文清單_row,b0700_row
    one_dissertationRow = []
    grad = my_record["502"]["a"]
    callnumber = my_record["084"]["a"] + " " + my_record["084"]["b"]
    barcode = my_record["035"]["a"]
    title = my_record["245"]["a"] + " " + my_record["245"]["b"]
    author = my_record["100"]["a"]
    edition = my_record["260"]["a"] + my_record["260"]["b"] + my_record["260"]["c"]
    one_dissertationRow.extend([grad, callnumber, barcode, title, author, edition])
    # =700 process
    one_0700Row = [""]*columns_0700_Len
    one_0700Row[columns_0700.index("authorC")] = author
    mentors = 0
    my_700s = my_record.get_fields("700")
    for my_700 in my_700s:
        name = my_700['a']
        mentors = mentors + 1
        if mentors < columns_0700.index("mentor4E")+1:  # only allow 4 mentors
            one_0700Row[mentors] = name
        else:
            print("Warning : more than 4 mentors.", author, "has more than", mentors//2, "mentors.")

    one_0700Row[columns_0700.index("author")] = author
    one_0700Row[columns_0700.index("titleOrig")] = title
    # NewTitle C546 process
    title_a = my_record["245"]["a"]
    c546_number, newTitle = getC546NewTitle(author, title_a)
    # 原本546 process
    if record["546"] is not None:
        field_546 = str(record["546"]["a"])
    else:
        field_546 = "無546"
    one_0700Row[columns_0700.index("原本546")] = field_546

    if not pd.isnull(newTitle):
        one_0700Row[columns_0700.index("titleNew")] = newTitle
    if not pd.isnull(c546_number):
        one_0700Row[columns_0700.index("c546")] = c546_number

    #
    return author, one_dissertationRow, one_0700Row

# -------------Step2.2: read already barcode updated file-----------------------
dissertation_list = []
b0700_list = []

with open(tempMRCfile, "rb") as fh:
    reader = MARCReader(fh)

    def addNoExistAuthorInMRC(authorFromXls):
        miss_dissertation_row = [""]*dissertation_len   # ['', '', '', '', 'tbu', '']
        miss_b0700_row = [""]*columns_0700_Len  # ['tbu','','','','','','','','','','','','tbu','','']
        miss_dissertation_row[columns_論文清單.index("作者")] = authorFromXls
        miss_b0700_row[columns_0700.index("authorC")] = authorFromXls
        miss_b0700_row[columns_0700.index("author")] = authorFromXls

        dissertation_list.append(miss_dissertation_row)
        b0700_list.append(miss_b0700_row)

    index_of_xls = 0
    authorCheck = "for double check purpose"

    for record in reader:
        authorFromMRC, dissertation_row, b0700_row = outputItems(record)
        found_flag = False
        while ((index_of_xls < len_of_barcode_xlsx) and (found_flag == False)):
            index_of_xls = index_of_xls + 1
            authorFromXls = item_name_dict_list[index_of_xls]["name"]
            if authorFromXls != authorFromMRC:
                addNoExistAuthorInMRC(authorFromXls)
            else:
                found_flag = True

        dissertation_list.append(dissertation_row)
        b0700_list.append(b0700_row)

    while (index_of_xls < len_of_barcode_xlsx):
        index_of_xls = index_of_xls + 1
        authorFromXls = item_name_dict_list[index_of_xls]["name"]
        addNoExistAuthorInMRC(authorFromXls)

# -------------Step2.3: output to excel files-----------------------
# (1)
df2 = pd.DataFrame(b0700_list, columns=columns_0700)
df2.index.name = "item"
df2.index += 1
df2.to_excel(output0700file)
# (2)
#print(a_list)
df1 = pd.DataFrame(dissertation_list, columns=columns_論文清單)
df1.index.name = "項次"
df1.index += 1
df1.to_excel(temp論文清單file)

# --------------------------------------------------------------------------------------
if err_list:
    print("\n")
    for e in err_list:
        print(e)
# --------------------------------------------------------------------------------------
