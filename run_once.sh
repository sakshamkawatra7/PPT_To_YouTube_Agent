#!/bin/bash

# Stop if a required command fails
set -e

# Move to the directory containing this script
cd "$(dirname "$0")"

# Activate the project's virtual environment
source .ytagent/bin/activate

# Process the next PowerPoint presentation
python ppt_youtube_agent.py --process-next
EOF