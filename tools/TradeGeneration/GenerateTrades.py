import os
import pandas as pd
import json
import random

def get_excel_data():
    """Reads data from the first Excel file found in the current directory and returns a DataFrame."""

    for file in os.listdir():
        if file.endswith(".xlsx") or file.endswith(".xls"):
            filename = file
            break
    else:
        raise FileNotFoundError("No Excel files found in the current directory.")

    try:
        df = pd.read_excel(filename, usecols=["Item_ID", "Profession", "Trade Level", "Buy Price", "Buy Amount", "Trade Type", "Sell Price", "Sell Amount", "Convert Item ID", "Convert Item Amount", "Max", "XP"], header=0)
        print(f"Successfully read Excel file: {filename}")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        raise

    return df

def transform_data(data, profession_prefix, trade_type_prefix):
    """Transforms the DataFrame into a dictionary of trades."""

    # Map numeric trade levels to text
    trade_levels = {
        1: "novice",
        2: "apprentice",
        3: "journeyman",
        4: "expert",
        5: "master"
    }

    # Initialize the main dictionary to store trades
    trades_data = {}

    # Handle missing values and convert data types
    data['Buy Price'] = data['Buy Price'].astype(int)
    data['Buy Amount'] = data['Buy Amount'].astype(int)
    data['Sell Price'] = data['Sell Price'].astype(int)
    data['Sell Amount'] = data['Sell Amount'].astype(int)

    # Iterate over each row in the DataFrame
    for _, row in data.iterrows():
        # Extract values from the current row
        item_id, profession, trade_level, buy_price, buy_amount, trade_type, sell_price, sell_amount, convert_item_id, convert_item_amount, max_uses, xp = row

        # Convert trade level to text and add prefix to profession
        trade_level = trade_levels[trade_level]
        profession = profession_prefix + ":" + profession

        # Initialize the profession in the `trades_data` dictionary if it doesn't exist
        if profession not in trades_data:
            trades_data[profession] = {"profession": profession, "trades": {level: [] for level in trade_levels.values()}}

        # Create trades based on the `trade_type`
        if trade_type == "Buy/Sell":
            # Create a buy trade
            buy_trade = {
                "type": f"{trade_type_prefix}:buy_item",
                "buy": {"item": "emerald", "count": buy_price},
                "reward": {"item": item_id, "count": buy_amount},
                "max_uses": max_uses if pd.notnull(max_uses) else random.randint(2, 8),
                "villager_experience": xp if pd.notnull(xp) else random.randint(2, 5)
            }
            trades_data[profession]["trades"][trade_level].append(buy_trade)

            # Create a sell trade
            sell_trade = {
                "type": f"{trade_type_prefix}:sell_item",
                "sell": {"item": "emerald", "count": sell_price},
                "priceIn": {"item": item_id, "count": sell_amount},
                "max_uses": max_uses if pd.notnull(max_uses) else random.randint(4, 12),
                "villager_experience": xp if pd.notnull(xp) else random.randint(2, 8)
            }
            trades_data[profession]["trades"][trade_level].append(sell_trade)
        elif trade_type == "Process":
            # Create a process trade
            process_trade = {
                "type": f"{trade_type_prefix}:process_item",
                "sell": {"item": "emerald", "count": sell_price},
                "priceIn": {"item": item_id, "count": buy_amount},
                "convertible": {"item": convert_item_id if pd.notnull(convert_item_id) else "dead_tube_coral", "count": int(convert_item_amount if pd.notnull(convert_item_amount) else 1)},
                "max_uses": max_uses if pd.notnull(max_uses) else random.randint(4, 12),
                "villager_experience": xp if pd.notnull(xp) else random.randint(2, 8)
            }
            trades_data[profession]["trades"][trade_level].append(process_trade)
        else:
            # Handle other trade types (currently assumes "Buy")
            if trade_type == "Buy":
                buy_trade = {
                    "type": f"{trade_type_prefix}:buy_item",
                    "buy": {"item": "emerald", "count": buy_price},
                    "reward": {"item": item_id, "count": buy_amount},
                    "max_uses": max_uses if pd.notnull(max_uses) else random.randint(2, 8),
                    "villager_experience": xp if pd.notnull(xp) else random.randint(2, 5)
                }
                trades_data[profession]["trades"][trade_level].append(buy_trade)

            if trade_type == "Sell":
                sell_trade = {
                    "type": f"{trade_type_prefix}:sell_item",
                    "sell": {"item": "emerald", "count": sell_price},
                    "priceIn": {"item": item_id, "count": sell_amount},
                    "max_uses": max_uses if pd.notnull(max_uses) else random.randint(4, 12),
                    "villager_experience": xp if pd.notnull(xp) else random.randint(2, 8)
                }
                trades_data[profession]["trades"][trade_level].append(sell_trade)

    return trades_data

def main():
    try:
        # Get the Excel data
        data = get_excel_data()

        # Get the profession prefix from the user
        profession_prefix = input("Enter the profession prefix (e.g., createengineers or delightfulchefs): ")

        # Determine the trade type prefix
        if profession_prefix.lower() == "minecraft":
            trade_type_prefix = input("Enter the trade type prefix (e.g., createengineers or delightfulchefs): ")
        else:
            trade_type_prefix = profession_prefix

        # Transform the data into a dictionary of trades
        trades_data = transform_data(data, profession_prefix, trade_type_prefix)
        print("Successfully transformed data.")

        # Write each profession's trades to a separate JSON file
        for profession, data in trades_data.items():
            file_name = f"{profession.replace(profession_prefix + ':', '')}.json"
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2)
                print(f"Successfully created JSON file: {file_name}")
            except Exception as e:
                print(f"Error creating JSON file: {e}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        input("Press Enter to exit...")

if __name__ == '__main__':
    main()