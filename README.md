# pymarc_thesis
Using Python "pymarc" to catalogue thesis.
1. Cataloguing original MARC: thesisMARC.py + namebarcode.xlsx + XXX.mrc + callnumber.xlsx --> tmp_thesisList.xlsx + 0700__infoRevise.xlsx + tmp_XXX.mrc
2. Cataloguing revised MARC: thesisMARXplus.py + tmp_thesisList.xlsx + 0700__infoRevise.xlsx + tmp_XXX.mrc --> final_XXX.mrc + final_thesisList.xlsx + doubleCheck.xlsx
3. Split MARC into dissertation & thesis: splitF502MRC.py + final_XXX.mrc --> finalXXX_008.8.mrc + finalXXX_008.9.mrc
