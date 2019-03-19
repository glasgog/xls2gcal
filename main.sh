#!/bin/bash

# Backup and download of xls shift file
# (url in xls_url text file)
mv Turni.xlsx Turni-`date +%y%m%d-%H%M%S`.xlsx.bak
wget -nv -i xls_url -O Turni.xlsx && python xls2gcal.py