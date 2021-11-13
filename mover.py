import datetime
import os
import win32com.client

###### O Path nao pode conter acentuacao ou cedilha

path = ''

###### Caso queira salvar em uma das pastas do usuario logado como desktop, documentos etc...

# path = os.path.expanduser('~/Desktop/arquivos')

################################

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

######## Caso Tenha mais de um arquivo PST ########

#email = outlook.Folders['guilherme@cairotecnologia.com.br']
#inbox = email.Folders[0]
############################################

inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
first = messages.GetFirst()

######## Aqui seria para criar uma pasta dentro do path pra organizar pelo Ano dos anexos ########
#if not os.path.exists(path+"/Certificado "+str(first.ReceivedTime.date().year)):
#    os.mkdir(path+"/Certificado "+str(first.ReceivedTime.date().year))

#path = path+"/Certificado "+str(first.ReceivedTime.date().year)

############################################

#attachments = first.Attachments

#attachment = attachments.Item(1)

for attachment in first.Attachments:
    attachment.SaveAsFile(os.path.join(path, str(attachment)))