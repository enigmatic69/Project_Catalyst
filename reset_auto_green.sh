#!/bin/bash

workdir="/usr/local/etc/project_catalyst"
filename="Project_Catalyst/ThisAddIn.cs"
bakup_filename="Project_Catalyst/ThisAddIn.cs.bak"


cp "$workdir/$bakup_filename" "$workdir/$filename"

# crontab -e
# 5 13 * * * /path/to/AutoGreenReset.sh
