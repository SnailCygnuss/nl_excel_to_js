#!/bin/sh

cd ../nlead_data_edit
git status
git branch
git add .
git commit -m "Update Data"
git push
open -a 'Google Chrome' https://github.com/SnailCygnuss/nlead_data_edit
echo "Opening Github on Google Chrome"