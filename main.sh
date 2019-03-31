#!/bin/bash

# Backup and download of xls shift file
# (url in xls_url text file)
rm data/*.xlsx.bak
mv data/Turni.xlsx data/Turni-`date +%y%m%d-%H%M%S`.xlsx.bak
echo "Download del file e lancio dell'applicazione in corso."
wget -nv -i "data/xls.url" -O "data/Turni.xlsx" && python xls2gcal.py
