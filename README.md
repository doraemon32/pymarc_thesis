# pymarc_thesis
Using Python "pymarc" to catalogue thesis.
1. Cataloguing an original MARC --> Revised MARC: myconfig.ini + thesisMARC.py + namebarcode.xlsm + 0221.mrc + 0221callnumber.xlsx --> tmp_thesisList.xlsx + 0700__infoRevise.xlsx + tmp_0221.mrc
2. Cataloguing a revised MARC --> Completed MARC: myconfig.ini + thesisMARXplus.py + tmp_thesisList.xlsx + 0700__infoRevise.xlsx + tmp_0221.mrc --> final_0221.mrc + callnumber.txt + final_thesisList.xlsx + doubleCheck.xlsx
3. Split MARC into dissertation & thesis: splitF502MRC.py + final_XXX.mrc --> finalXXX_008.8.mrc + finalXXX_008.9.mrc
MARCtoCallnumber.py will be used when you only need to create a callnumber.txt. *You just need to input a completed MARC and a **namebarcode.xlsm** **(optional)**
