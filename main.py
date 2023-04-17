import win32com.client as win32

# conectar código ao outlook
outlook = win32.Dispatch('outlook.application')
# Criar um email
email = outlook.CreateItem(0)
# detalhes do email
email.To = 'mvictoriasilva114@gmail.com'#Destino
email.Subject = 'Testando código'#Assunto
email.HTMLBody = '''
<p>Olá Victória, estou testando meu código!</p>

<p>De acordo com o python, o envio de email solicitado funcionou!</p>

<p>Atenciosamente,</p>
<p>Diogo Santana</p>
'''
anexo = "C:\\Users\\diogo\\Desktop\\curriculo 2022\\CV_DiogoReis_2023_pt-Br.pdf"
email.Attachments.Add(anexo)


email.Send()
print('Email enviado!')