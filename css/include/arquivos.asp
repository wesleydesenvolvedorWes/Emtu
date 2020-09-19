<!--#include Virtual="/aplicativos/limpaStringRequest.asp"-->

<%

if banco = "" then 

	Set  banco = Server.CreateObject("ADODB.Connection")
	with banco

	' -----------------------------------------
	'      WEBSERVER - BANCO PRODUÇÃO 
	' -----------------------------------------
	'    NÃO DEVE SER USADO POIS CONVERSA COM BENNER/GESTEC/SITE
	'    .connectionString = "Provider=SQLOLEDB.1;DRIVER={System.Data.SqlClient.SqlConnection};SERVER=web;UID=sa;PWD=vitorsofia;DATABASE=dbemtu"        

	' -----------------------------------------
	'      WEBSERVER - BANCO PARA TESTES
	' -----------------------------------------
	'      LIVRE ACESSO
	    .connectionString = "Provider=SQLOLEDB.1;DRIVER={System.Data.SqlClient.SqlConnection};SERVER=web;UID=sa;PWD=vitorsofia;DATABASE=dbemtu_teste"        
		 
	      .Open
	end with
	if err.number <> 0 then 
	   response.redirect "/login.html"
	   response.end
	end if       

end if 


 %>