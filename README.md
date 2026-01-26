# Laptop Inventory Manager - README

## Overview
This PowerShell application provides a graphical interface for managing laptop inventory and comparing it against CSV files.

## Features

### Tab 1: CSV Import & Compare
- **Browse and Load CSV**: Import CSV files containing laptop data
- **View CSV Data**: Display all records from the imported CSV file
- **Compare with Inventory**: Cross-reference CSV data with your in-stock inventory
- **Comparison Results**: See which laptops are in your inventory and which are not

### Tab 2: Manage Inventory
- **Add Laptops**: Add new laptops to your inventory with:
  - Serial Number (required)
  - Model (optional)
  - Location (dropdown: Denver, Greenwood Village, Salt Lake City)
  - Loaner status (checkbox)
- **View Current Inventory**: See all laptops currently in your inventory
- **Remove Laptops**: Select and remove laptops from inventory
- **Refresh**: Reload inventory data

## How to Use

### Running the Application
1. Open PowerShell
2. Navigate to the folder containing `InventoryManager.ps1`
3. Run: `.\InventoryManager.ps1`

### Adding Laptops to Inventory
1. Go to the "Manage Inventory" tab
2. Enter the serial number (required)
3. Optionally enter the model
4. Select a location from the dropdown
5. Check "Loaner Laptop" if applicable
6. Click "Add to Inventory"

### Importing and Comparing CSV
1. Go to the "CSV Import & Compare" tab
2. Click "Browse..." and select your CSV file
3. Click "Load CSV" to view the data
4. Click "Compare with Inventory" to see which laptops match your inventory

### Removing Laptops
1. Go to the "Manage Inventory" tab
2. Select a laptop from the list
3. Click "Remove Selected"
4. Confirm the removal

## Data Storage
- Inventory data is stored in `inventory.json` in the same folder as the script
- The file is automatically created on first run
- Data persists between sessions

## CSV Format
The application expects CSV files with at least a `serial_number` column. Additional columns like `model`, `install_status`, and `stockroom` will be displayed in the comparison results.

## Requirements
- Windows PowerShell 5.1 or later
- Windows operating system (for Windows Forms)
