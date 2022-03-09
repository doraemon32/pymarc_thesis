from pymarc import MARCReader

inputfile = "final_0221_fc_2022022311.mrc"
outputfile1 = "0221fc1_0088.mrc"
outputfile2 = "0221fc1_0089.mrc"

counter88 = 0
counter89 = 0
my_marc_records_0088 = []
my_marc_records_0089 = []
with open(inputfile, "rb") as fh:
    reader = MARCReader(fh)
    counter = 0

    for record in reader:
        grad = record["502"]["a"][0:2]  #e.g. "碩士"
        if grad == "碩士":
            #grad_number = "008.8"
            counter88 = counter88 + 1
            my_marc_records_0088.append(record)
        else:
            #grad_number = "008.9"
            counter89 = counter89 + 1
            my_marc_records_0089.append(record)
        counter = counter + 1
print("0088  : ",counter88)
print("0089  : ",counter89)
print("Total : ",counter)
#------------------------------------------------------------
if counter88 > 0:
    with open(outputfile1, "wb") as out:
        for my_record in my_marc_records_0088:
            ### and write each record to it
            out.write(my_record.as_marc())

if counter89 > 0:
    with open(outputfile2, "wb") as out:
        for my_record in my_marc_records_0089:
            ### and write each record to it
            out.write(my_record.as_marc())
