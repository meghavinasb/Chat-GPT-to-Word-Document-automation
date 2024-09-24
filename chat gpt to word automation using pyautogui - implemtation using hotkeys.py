import pyautogui as p
import time as t

# Using prompt for user input
a = p.prompt("What is your question/prompt?")

# Open Chrome using the Windows search
p.hotkey('win', 's')  # Opens Windows search
t.sleep(1)
p.typewrite('chrome')
p.press('enter')
t.sleep(2)

# Typing the URL in the address bar
url = "https://chat.openai.com/"
p.hotkey('ctrl', 'l')  # Focus on the address bar
p.typewrite(url)
p.press('enter')

# Adding a delay to let the page load fully before interacting with it
t.sleep(7)

# Navigating to the search bar using hotkeys (assuming TAB will reach the input field)
p.press('tab', presses=10)  # Adjust number of presses if needed to reach the input field
t.sleep(1)

# Enter the user's prompt
p.typewrite(a)
p.press('enter')

# Wait for the response to be generated (adjust the sleep time as needed)
t.sleep(20)

# Copy the ChatGPT response using hotkeys
p.hotkey('ctrl', 'a')  # Select all
p.hotkey('ctrl', 'c')  # Copy

# Open Microsoft Word using the Windows search
p.hotkey('win', 's')
t.sleep(1)
p.typewrite('word')
p.press('enter')
t.sleep(2)

# In case Microsoft Word opens with an activation wizard, close it using hotkeys
# (you may adjust based on what appears, this step is left generic)
# You could also add some logic to check for the activation window, but this is a placeholder
t.sleep(2)

# Paste the ChatGPT response into Word
p.hotkey('ctrl', 'v')
t.sleep(1)

# Save the document
p.hotkey('ctrl', 's')
t.sleep(3)

# Set filename as user's prompt
p.typewrite(a)
t.sleep(1)

p.press('enter')

# Close Microsoft Word
p.hotkey('alt', 'f4')
