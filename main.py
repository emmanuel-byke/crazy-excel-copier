import pandas as pd
import pyperclip
import pyautogui
import sys

def main():
    if len(sys.argv) != 2:
        print("Usage: python excel_paster.py <path_to_excel_file>")
        return

    file_path = sys.argv[1]

    try:
        # Read Excel file (assuming no header row)
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Collect all cell values in row-major order
    cell_values = []
    for _, row in df.iterrows():
        for cell in row:
            cell_values.append(str(cell) if pd.notna(cell) else '')

    print(f"Loaded {len(cell_values)} cells. Starting paste sequence...")
    print("-----------------------------------------------------------")

    for i, value in enumerate(cell_values):
        # Copy value to clipboard
        pyperclip.copy(value)
        
        print(f"Cell {i+1} ready: {value[:50]}...")  # Show preview
        
        input("1. Focus target text box\n2. Press Enter to paste...")
        
        # Send paste command
        pyautogui.hotkey('ctrl', 'v')
        
        print("Pasted! Waiting for next cell...\n")

    print("All cells processed!")

if __name__ == "__main__":
    main()