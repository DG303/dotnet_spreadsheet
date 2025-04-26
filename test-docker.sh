#!/bin/bash

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Test file name
TEST_FILE="test.xlsx"
TEST_VALUE="Hello, Docker!"

echo "Starting Docker test for Spreadsheet Editor..."

# Step 1: Create a new Excel file
echo -e "\n${GREEN}Step 1: Creating new Excel file...${NC}"
docker-compose run --rm spreadsheet-editor --file /app/data/$TEST_FILE --new
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to create Excel file${NC}"
    exit 1
fi
echo -e "${GREEN}✓ Excel file created successfully${NC}"

# Step 2: Write a value to cell A1
echo -e "\n${GREEN}Step 2: Writing value to cell A1...${NC}"
docker-compose run --rm spreadsheet-editor --file /app/data/$TEST_FILE --cell A1 --value "$TEST_VALUE" --write
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to write to cell${NC}"
    exit 1
fi
echo -e "${GREEN}✓ Value written successfully${NC}"

# Step 3: Read the value from cell A1
echo -e "\n${GREEN}Step 3: Reading value from cell A1...${NC}"
OUTPUT=$(docker-compose run --rm spreadsheet-editor --file /app/data/$TEST_FILE --cell A1 --read)
if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to read from cell${NC}"
    exit 1
fi

# Verify the read value matches what we wrote
if [[ $OUTPUT == *"$TEST_VALUE"* ]]; then
    echo -e "${GREEN}✓ Value read successfully and matches what was written${NC}"
    echo -e "Read value: $OUTPUT"
else
    echo -e "${RED}✗ Read value does not match what was written${NC}"
    echo -e "Expected: $TEST_VALUE"
    echo -e "Got: $OUTPUT"
    exit 1
fi

echo -e "\n${GREEN}All tests passed successfully!${NC}"

# Clean up
echo -e "\nCleaning up test file..."
rm -f ./data/$TEST_FILE 