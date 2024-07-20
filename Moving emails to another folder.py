# Moving emails to another folder

import win32com.client

Outlook = win32com.client.Dispatch("Outlook.Application")
namespace = Outlook.GetNamespace("MAPI")

carpetas = namespace.Folders

segunda_cuenta = carpetas[1]

bandeja_entrada_segunda_cuenta = segunda_cuenta.Folders("Bandeja de entrada")

counter = 0
for i in bandeja_entrada_segunda_cuenta.Items:
    if counter == 5000:
        break
    else:
        if "iberempleos.es" in i.SenderEmailAddress:
            try:
                i.Move(segunda_cuenta.Folders("Iberempleos"))
                counter = counter + 1
            except Exception as e:
                continue