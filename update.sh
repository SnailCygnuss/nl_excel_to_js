#!/bin/sh

GIT_DIR="nlead_data"
F1="research_calls.js"
F2="research_conf.js"
F3="seminar_webinar.js"
F4="special_issues.js"
F5="updated_date.js"

if [ ! -d "$GIT_DIR" ]; then
    echo "Cloning git folder"
    git clone "https://github.com/snailCygnuss2/nlead_data.git"
fi


echo "Updating Webpages"
echo "Checking Excel Files"
source venv/bin/activate
venv/bin/python3 create_html.py
deactivate

# Move files to Git folder
cp $F1 $GIT_DIR/data/
cp $F2 $GIT_DIR/data/
cp $F3 $GIT_DIR/data/
cp $F4 $GIT_DIR/data/
cp $F5 $GIT_DIR/data/

git status --porcelain | egrep " M data/"