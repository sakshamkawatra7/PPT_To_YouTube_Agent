#!/bin/bash

cd "$(dirname "$0")"

source .ytagent/bin/activate

python ppt_youtube_agent.py --auth

status=$?

echo
read -r -p "Press Enter to close this window..."

exit "$status"
EOF