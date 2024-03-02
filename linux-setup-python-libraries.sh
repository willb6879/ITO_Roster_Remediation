#!/usr/bin/bash
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
RESET='\033[0m'

# Checks status of install command
check_status(){
    if [ $? -eq 0 ]; then
        echo -e "$GREEN Success: $1\n"
    else
        echo -e "$RED Error: $1 failed installing\n"
        exit 1
    fi
}

# Install necessary libraries for 'user-list-csv-tool.py'
echo -e "$RESET\n--- Installing pip (Python package installer) using apt ---"
sudo apt update
sudo apt install pip -y
check_status "'apt install pip'"

echo -e "$RESET\n--- Installing pandas using pip ---"
pip install pandas
check_status "'apt install pandas'"

echo -e "$RESET\n--- Installing openpyxl using pip ---"
pip install openpyxl 
check_status "'pip install openpyxl'"