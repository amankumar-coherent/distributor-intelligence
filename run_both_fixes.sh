#!/bin/bash
# GLP-1 first, then Wearable Tech on the same corrected logic.
cd "C:/Users/aman.kumar/Downloads/vendor-intel"
PY=".venv/Scripts/python.exe"
D="C:/Users/aman.kumar/Desktop/DI"
$PY fix_glp1_columns.py --file "$D/GLP1_Receptor_Agonist_Distributors.xlsx"  > logs_glp1_fix.txt     2>&1
$PY fix_glp1_columns.py --file "$D/Wearable_Tech_Distributors.xlsx"          > logs_wearable_fix.txt 2>&1
echo "BOTH FILES COMPLETE"
