#!/bin/bash

set -e

cd "$(dirname "$0")"

source .ytagent/bin/activate

python -m uvicorn server:app --host 127.0.0.1 --port 8000
EOF