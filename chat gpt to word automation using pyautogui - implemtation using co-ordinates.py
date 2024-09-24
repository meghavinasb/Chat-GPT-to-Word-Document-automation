import pyautogui as p
import time as t

# Using prompt for user input
a = p.prompt("What is your question/prompt?")

# Pressing windows search bar and typing 'chrome' to search for Chrome App
p.press('win')
t.sleep(1)
p.typewrite('chrome')
p.press('enter')
t.sleep(2)
x=871
y=615
p.click(x,y)

# Typing the URL in the address bar
url = "https://chat.openai.com/"
p.typewrite(url)
p.press('enter')

# Adding a delay to let the page load fully before typing the channel name
t.sleep(7)

# Pressing the search bar in chatgpt
z = 766
q = 971
p.click(z,q)
t.sleep(1)

# Prompting the users prompt
p.typewrite(a)
p.press('enter')

# Copying the answer
t.sleep(20)
r = 794
s = 889
p.click(r,s)


# Opening Microsoft Word 
p.press('win')
t.sleep(1)
p.typewrite('Microsoft Word')
p.press('enter')
t.sleep(2)

# Closing microsoft office activation wizard
p.click(1161,711)


# Pating the answer given by chat gpt
p.hotkey('ctrl','v')
t.sleep(1)

# Saving the file
p.hotkey('ctrl','s')
t.sleep(3)

filename = a # Saving the file as the prompt given by the user
p.typewrite(filename)
t.sleep(1)

p.press('enter')

#Closing the File Explorer
p.hotkey('alt','f4')
