from pymarc import MARCReader
from configparser import ConfigParser
from datetime import datetime
import pandas as pd
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# -MARCtoCallnumber.py get configure file                                             --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# define filename by datetime
now = datetime.now()
yymmddhh = now.strftime("%Y%m%d%H")
# Read tocallnumber.ini file
config_object = ConfigParser()
config_object.read("tocallnumber.ini")
# Get and display section
myinfo = config_object["DEFAULT"]
input_fileinfo = config_object["INPUTFILE"]
output_fileinfo = config_object["OUTPUTFILE"]
# [DEFAULT] section
needAuthorfile = myinfo["needauthorfile"]
# [INPUTFILE] section
origMRCfilename = input_fileinfo["origmrcfile"]
if needAuthorfile.lower() == "yes":
    authorfilename = input_fileinfo["authorfile"]
else:
    authorfilename = "no_authorfile"
# add datetime for outputfile name
origMRCfilename_list = origMRCfilename.split(".")
# [OUTPUTPUTFILE] section
if needAuthorfile.lower() == "yes":
    finalMRCfile = origMRCfilename_list[0] + "_fc_" + yymmddhh + ".mrc"
else:
    finalMRCfile = ""
output_fileinfo["mrcfile"] = finalMRCfile
論文清單file = origMRCfilename_list[0] + "_論文清單_" + yymmddhh + ".xlsx"
output_fileinfo["論文清單file"] = 論文清單file
callNumberfile = origMRCfilename_list[0] + "_callnumber_" + yymmddhh + ".txt"
output_fileinfo["callnumberfile"] = callNumberfile
# Write changes back to ini file
with open("tocallnumber.ini", "w") as conf:
    config_object.write(conf)

# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
# - use authorfile for insequence purpose or not                                      --
# --------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------
err_list = []
my_marc_records = []

if authorfilename != "no_authorfile":
    # --------------------------------------------------------------------------------------

    # - retrieve excel to list
    # --------------------------------------------------------------------------------------
    authorfilename_Df = pd.read_excel(authorfilename, sheet_name=0, header=0)  # header=None means no header row
    authorfilename_Df.columns = ["item", "name", "barcode"]
    len_of_authorfile_xlsx = len(authorfilename_Df.index)
    item_name_dict_list = authorfilename_Df.set_index('name').T.to_dict(orient='dict')

    my_marc_sequence = []
    my_marc_records_tmp = []

    with open(origMRCfilename, "rb") as fh:
        reader = MARCReader(fh)
        counter = 1
        for record in reader:
            author = record["100"]["a"]
            print(counter, author, record["245"]["a"])   # title
            if author in item_name_dict_list:    # get['項次'] based on '作者'
                itemOrder_inXls = item_name_dict_list[author]['item']
            else:
                my_err = "**Error::mrc author "+author+" NotFound in the excel file.**"
                err_list.append(my_err)
                #print(my_err)
            counter = counter + 1
            # itemOrder_inXls for later in-sequence purpose
            my_marc_sequence.append(itemOrder_inXls)
            my_marc_records_tmp.append(record)

    for author_name in item_name_dict_list:
        index_of_xls = item_name_dict_list[author_name]['item']
        if index_of_xls not in my_marc_sequence:
            my_err = "**Error Index "+str(index_of_xls)+":"+author_name+" does not exist in the mrc file.**"
            err_list.append(my_err)
            print('\nList the name sequence number from~', authorfilename, 'file,dependent on~', origMRCfilename, 'file sequence.\n', my_marc_sequence)
            #print(my_err)
        else:
            tmp_record = my_marc_records_tmp[my_marc_sequence.index(index_of_xls)]
            my_marc_records.append(tmp_record)
    # ----------------------------------------------output final mrcfile------------
    with open(finalMRCfile, "wb") as out:
        for record in my_marc_records:
            # -- and write each record to it
            out.write(record.as_marc())

else:  # authorfilename == "no_authorfile"

    with open(origMRCfilename, "rb") as fh:
        reader = MARCReader(fh)
        for record in reader:
            my_marc_records.append(record)
            print(record["100"]["a"], record["245"]["a"])   # title


# --------------------------------------------------------------------------------------
#  output Program :
# --------------------------------------------------------------------------------------

# ----------------------------------------------procedure file ------------
def outputItems(record):
    a_dissertation_list = []

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
    elif authorfilename != "no_authorfile":
        barcode = str(item_name_dict_list[author]["barcode"])
    else:
        my_err = "**Warning::mrc author "+author+" barcode not exists.**"
        err_list.append(my_err)
        #print(my_err)
    a_dissertation_list.extend([grad, callnumber, barcode, title, author, edition])
    # above for 論文清單
    return a_dissertation_list

# ------------------------------------------prepare output process------------
dissertation_list = []
prtCallnumber = ""

for record in my_marc_records:
    # 論文清單 process
    a_dissertation_list = outputItems(record)
    dissertation_list.append(a_dissertation_list)

    # callNumber process
    grad_number_R = record["084"]["a"]
    fourCorner_list = record["084"]["b"].split(" ")
    # follow the spec format
    prtCallnumber = prtCallnumber+grad_number_R[:1]+"\n"+grad_number_R[-5:]+"\n"+fourCorner_list[0]+"\n"+fourCorner_list[1]+"\n\n\n\n\n\n\n\n"
#print(prtCallnumber)
# ----------------------------------------------output 2------------
# (1)
with open(callNumberfile, 'w') as f:  # callNumber output process
    f.write(prtCallnumber)
# (2)
columns_dissertation = ['系所', '索書號', '條碼', '書名', '作者', '出版項']
df1 = pd.DataFrame(dissertation_list, columns=columns_dissertation)
df1.index.name = "項次"
df1.index += 1
df1.to_excel(論文清單file)
# --------------------------------------------------------------------------------------
if err_list:
    print("\n")
    for e in err_list:
        print(e)
# --------------------------------------------------------------------------------------

