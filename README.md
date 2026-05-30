# Grocery Inventory Management System

A simple Python program to manage your home grocery inventory with the ability to export to Excel spreadsheets.

## Features

✨ **Core Functionality:**
- 📝 Add items to your grocery inventory
- 🗑️ Remove items from the inventory
- ✏️ Update item quantities and expiry dates
- 📋 View all items in a formatted table
- 🔍 Filter items by category
- 📊 Export inventory to Excel with formatted headers and styling

## Installation

### Prerequisites
- Python 3.6+
- pip (Python package manager)

### Setup

1. Clone or navigate to the HelloAI repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the program:
```bash
python grocery_inventory.py
```

### Menu Options

1. **View Inventory** - Display all items in a formatted table
2. **Add Item** - Add a new grocery item with quantity, unit, and category
3. **Remove Item** - Delete an item from inventory by ID
4. **Update Item** - Modify quantity or expiry date
5. **View by Category** - Filter and display items by category
6. **Export to Excel** - Save your inventory to an Excel file
7. **Exit** - Close the program

## Example Workflow

```
Add items:
- Apples, 5 kg, Fruits
- Milk, 2 liters, Dairy
- Carrots, 3 kg, Vegetables

Update an item:
- Change Apples quantity to 3 kg

Export to Excel:
- Save as "my_grocery_list.xlsx"
```

## Data Storage

- Inventory data is automatically saved to `inventory.json`
- Each item includes: ID, name, quantity, unit, category, expiry date, and date added
- Data persists between program runs

## Excel Export

The exported Excel file includes:
- Formatted header row with blue background and white text
- Columns: ID, Item Name, Quantity, Unit, Category, Expiry Date, Added Date
- Auto-adjusted column widths for readability

## File Structure

```
HelloAI/
├── grocery_inventory.py    # Main program
├── requirements.txt         # Python dependencies
├── README.md               # This file
├── inventory.json          # Data file (auto-created)
└── grocery_inventory.xlsx  # Excel export (generated)
```

## Example Categories

- Fruits
- Vegetables
- Dairy
- Grains
- Meat/Poultry
- Pantry
- Beverages
- Frozen Foods

Feel free to add your own categories!

## License

This project is open source and available for personal use.
