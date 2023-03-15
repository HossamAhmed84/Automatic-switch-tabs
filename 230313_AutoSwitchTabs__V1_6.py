import webbrowser
import time
import pyautogui

# Define the time interval between website switches (in seconds)
switch_interval = 180

# Define the screen sizes and their centers
#laptop 1920x1080
#Screen 1366x768
#screen1 = (1920, 0, 1920, 1080)
screen1 = (1366, 0, 1366, 768)
#screen2 = (0, 0, 1920, 1080)
screen2 = (0, 0, 1366, 768)
#center1 = (960, 540)
center1 = (6, 70)
#center2 = (2880, 540)
center2 = (-1360, 70)

# Define the keyboard shortcut to switch between windows
switch_windows = ['alt', 'tab']

# Define the keyboard shortcut to switch between tabs
switch_tabs = ['ctrl', 'tab']



# Define a function to refresh the website lists
def refresh_website_lists():
    # Code to refresh the website lists goes here
    pass

# Specify the path to your Google Chrome executable
chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'
msEdge_path ="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
# Register the browser with webbrowser
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
webbrowser.register('msEdge', None, webbrowser.BackgroundBrowser(msEdge_path))
file_path = "Website_Lists.txt" # replace with your file path
file_path1 = "Website_Lists1.txt" # replace with your file path
website_lists=[]
website_lists1=[]
current_tabs = []
current_tabs1 = []
with open(file_path, "r") as f:
    data = f.read().splitlines()
with open(file_path1, "r") as f:
    data1 = f.read().splitlines()
for row in data:
    row = 'http://'+row+'/ecsweb/index.php?module=staterec'
    #print (row) 
    website_lists.append(row)
for row in data1:
    row = 'http://'+row+'/ecsweb/index.php?module=staterec'
    #print (row) 
    website_lists1.append(row)

#print (website_lists)
open_sites1 =0
full_screen1 = 0
open_sites =0
full_screen = 0
i=0
j=0
# Loop indefinitely
while True:
    #Chrome loop   
    if open_sites ==0:
        for website in website_lists:
            # Open the website in Google Chrome in full-screen mode
            tab = webbrowser.get('chrome').open_new_tab(website)
            # Open the website in a new window
            #webbrowser.get('chrome').open_new(website)
            current_tabs.append(tab)
            time.sleep(20)
            
        open_sites=1
        
    if full_screen ==0:
        # Press F11 to enter full-screen mode
        pyautogui.hotkey('f11')  
        full_screen =1
   

    #MS Edge loop    
    if open_sites1 ==0:
        for website in website_lists1:
           # Open the website in MS Edge in full-screen mode
            tab1 = webbrowser.get('msEdge').open_new_tab(website)
            # Open the website in a new window
            #webbrowser.get('msEdge').open_new(website)
            current_tabs1.append(tab)
            time.sleep(20)       
        open_sites1=1  
        
        
    if full_screen1 ==0:
        # Press F11 to enter full-screen mode
        pyautogui.hotkey('f11')  
        full_screen1 =1 
        
    # Bring the first screen to the foreground and activate Chrome
    pyautogui.moveTo(center1)
    pyautogui.click()
    
    time.sleep(5)
    pyautogui.click()
    
    time.sleep(5)
    pyautogui.hotkey(*switch_tabs)
     
    
    # Bring the second screen to the foreground and activate Edge
    pyautogui.moveTo(center2)
    pyautogui.click()
    
    time.sleep(5)
    pyautogui.click()
    
    time.sleep(5)  
    pyautogui.hotkey(*switch_tabs)    
    
    
    # Wait for the switch interval
    time.sleep(switch_interval)
    
    # Refresh the website lists
    refresh_website_lists()