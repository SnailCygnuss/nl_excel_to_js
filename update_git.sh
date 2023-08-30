#!/bin/zsh

cd ../nlead_data_edit
git status
vared -p "Press Enter to continue" -c temp
git checkout data_update
vared -p "Press Enter to continue" -c temp
git add .
vared -p "Press Enter to continue" -c temp
git commit -m "Update Data"
git pull
git push
open -a 'Google Chrome' https://github.com/SnailCygnuss/nlead_data_edit
echo "Opening Github on Google Chrome"