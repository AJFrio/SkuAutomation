import pyautogui as pg
import keyboard
import time

def save_cursor_position():
    # Get current mouse position
    x, y = pg.position()
    # Format as [x,y],
    formatted_position = f"[{x},{y}]"
    
    # Append to data.txt
    with open("data.txt", "a") as file:
        file.write(formatted_position + ":")
    
    print(f"Saved position: {formatted_position}")

# Register the End key to trigger saving cursor position
keyboard.add_hotkey('end', save_cursor_position)

print("Press End key to save cursor position. Press Ctrl+C to exit.")
# Keep the script running
try:
    while True:
        time.sleep(0.1)
except KeyboardInterrupt:
    print("\nExiting program")
