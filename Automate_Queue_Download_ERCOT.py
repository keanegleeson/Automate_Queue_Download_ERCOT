#Automate Queue Download ERCOT
#Make sure spreadsheets are not left open when this runs or it will not work


import pyautogui as pag
import webbrowser
import time
# pag.typewrite('Hello world!', interval=0.25)Hello world!
# print(pag.size())

#

#1) Go to ISO site
webbrowser.open_new("http://mis.ercot.com/misapp/GetReports.do?reportTypeId=15933&reportTitle=GIS%20Report&showHTMLView=&mimicKey")

#2) Click download
time.sleep(2)
pag.click(x=749,y=174)

#3) Move file to desired location on G drive ->  G:\Advanced Visualization\Staging\Common\NEO\Interconnection Queues
#3.1a Click the download
time.sleep(5)
pag.click(x=99,y=1008)

#3.1b Have the program wait a few seconds
time.sleep(5)

#3.1c Click on the excel window
pag.click(x=592,y=15)

#3.2 Open Save as dialoague (f12)
pag.press('f12')
#3.3 Click filepath
pag.click(x=512,y=45)

#3.4 Eneter in file location: G:\Advanced Visualization\Staging\Common\NEO\Interconnection Queues
pag.typewrite('G:\\Advanced Visualization\\Staging\\Common\\NEO\\Interconnection Queues')

#3.5 Press enter

pag.press('enter')
#3.6 Click Filename
pag.click(x=687,y=402)
#3.7 Enter app filename ERCOT1
pag.typewrite('ERCOT1')
#3.8 Click save
pag.click(x=772,y=538)

#3.9 Click yes on replace
time.sleep(2)
pag.click(x=1014,y=526)

#4 Close Excel file
time.sleep(6)
pag.click(x=1904,y=11)
