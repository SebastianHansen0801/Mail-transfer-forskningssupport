from win32com.client import Dispatch
import os
import re


f = '1135 - Indeklima / Indoor Environment'
f1 = 'Overført til server: 1130 - Energi og installationer / Energy and Services'

# opening application
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# opening inbox
for i in range(10):
    try:
        inbox = outlook.Folders["BYG-Forskningssupport"].Folders["Indbakke"].Folders[i]
        print(str(i) + ': ' + inbox.Name)
    except:
        print(str(i) + ': ' + 'none')
