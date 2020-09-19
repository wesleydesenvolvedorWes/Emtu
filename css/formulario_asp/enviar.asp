<%
Server.ScriptTimeOut = 60

'recebendo as variáveis
conteudo =Request("conteudo")
nome =Request("nome")
email =Request("email")
email2 = "Seu email vem aqui"

content=conteudo

Set objMail = Server.CreateObject("CDONTS.NewMail")
objMail.From = email
objMail.To = email2
objMail.Subject = nome
objMail.Body = content
objMail.BodyFormat = 0
objMail.MailFormat = 0
objMail.Send

Set objMail = Nothing

response.write "Sua mensagem foi 
enviada com Sucesso - Eviaremos uma 
resposta assim que possivel - Você pode 
fechar essa janela "

Set objMail = Nothing 
%>