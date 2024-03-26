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
echo -e "$RESET\n--- Installing python3 and pip (Python package installer) using choco ---"
choco install python3
check_status "'choco install python3'"

echo -e "$RESET\n--- Installing pandas using pip ---"
pip install pandas
check_status "'pip install pandas'"