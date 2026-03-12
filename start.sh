#!/bin/bash
cd "$(dirname "$0")"
echo ""
echo "  GrUE Sustainability Report Generator"
echo "  Starting server..."
echo ""

# Copy .env.example if .env doesn't exist
if [ ! -f .env ]; then
  cp .env.example .env
  echo "  Created .env file. Add your ANTHROPIC_API_KEY to .env (optional — you can also enter it in the app)."
  echo ""
fi

# Load .env
if [ -f .env ]; then
  export $(grep -v '^#' .env | xargs) 2>/dev/null
fi

python3 server.py
