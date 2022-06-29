import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)

email.To = "arturbos2021@gmail.com"
email.Subject = "E-mail automatizado"
email.HTMLBody = f"""
<p>Olá Artur Bomtempo Sales, estou te enviando um email referente ao faturamento da loja enquanto uso Python.</p>

<p>O faturamento da loja foi de R$1.500</p>
<p>Nós vendemos 10 produtos</p>
<p>O ticket médio foi de R$150</p>

<p>Tenha um bom dia!</p>
"""

email.Send()
print("E-mail enviado")