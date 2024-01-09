#!/bin/bash

workdir="/usr/local/etc/project_catalyst"
filename="Project_Catalyst/ThisAddIn.cs"

current_time=$(date)

echo "// $current_time" >> "$workdir/$filename"

cd $workdir

git add -A

git commit -m "A commit a day keeps your girlfriend away. ($current_time)"

git push 

# crontab -e
# 5 14 * * * /path/to/AutoGreen.sh
