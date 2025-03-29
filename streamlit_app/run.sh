#!/bin/bash

# Check if the API environment variable is set
if [ -z "$API" ]; then
    echo "Error: API environment variable is not set."
    echo "Please set it with: export API=your_openai_api_key"
    exit 1
fi

# Run the Streamlit app
streamlit run app.py 