import pandas as pd
import pyperclip
import sys
import keyboard

def main():
    if len(sys.argv) != 2:
        print("Usage: python excel_paster.py <path_to_excel_file>")
        return

    file_path = sys.argv[1]

    try:
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    cell_values = []
    for _, row in df.iterrows():
        for cell in row:
            cell_values.append(str(cell) if pd.notna(cell) else '')

    print(f"Loaded {len(cell_values)} cells. Starting paste sequence...")
    print("-----------------------------------------------------------")
    print("HOW TO USE:")
    print("1. Focus target text field after this message")
    print("2. First cell is automatically copied to clipboard")
    print("3. Paste with Ctrl+V")
    print("4. Press Ctrl+Right to advance to next cell")
    print("5. Press Esc to exit at any time\n")

    current_index = 0
    pyperclip.copy(cell_values[current_index])
    print(f"Cell {current_index+1} ready: {cell_values[current_index][:50]}...")

    while current_index < len(cell_values):
        # Wait for either advance command or exit
        try:
            event = keyboard.read_event() 
            if event.event_type == keyboard.KEY_DOWN:
                if event.name == 'esc':
                    print("\nExiting...")
                    break
                
                # Ctrl+Right arrow detection
                if keyboard.is_pressed('ctrl') and event.name.lower() == 'v':
                    current_index += 1
                    if current_index < len(cell_values):
                        pyperclip.copy(cell_values[current_index])
                        print(f"\nCell {current_index+1} ready: {cell_values[current_index][:50]}...")
                        # Auto-paste if desired (uncomment next line)
                        # pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            print(f"\nError: {e}")
            break

    print("\nPasting session ended.")
if __name__ == "__main__":
    main()