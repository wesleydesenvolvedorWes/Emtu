
<%

'BIBLIOTECA DE FUNCOES DE VALIDACAO'

'FUNCAO QUE FILTRA CARACTERES MALICIOSOS'
'-------------------------------------'
function SafeVariavelSql(valor)
'-------------------------------------'
	'if valor <> "" then
	''	SafeVariavelSql = replace(valor,"'","`")
	'end if

     if valor <> "" then
          'SUBSTITUICAO DE CARACTERES/PALAVRAS PROIBIDAS'
          valor = Replace(valor,"'","`")
          valor = Replace(valor, "|", "!")
          valor = Replace(valor, "--", "..")
          valor = Replace(valor, "/*", "")
          valor = Replace(valor, "*/", "")
          valor = Replace(valor, "@@", "@-@")
          valor = Replace(valor, "sysobjects", "", 1, -1, 1)   'case insensitive'
          valor = Replace(valor, "syscolumns", "", 1, -1, 1)
          valor = Replace(valor, "sysdatabases", "", 1, -1, 1)
          valor = Replace(valor, "execute ", "executa", 1, -1, 1)
          valor = Replace(valor, "exec ", "executa ", 1, -1, 1)
          valor = Replace(valor, "create ", "criar ", 1, -1, 1)
          valor = Replace(valor, "alter table", "alterar ", 1, -1, 1)
          valor = Replace(valor, "drop ", "excluir ", 1, -1, 1)
          valor = Replace(valor, "insert ", "insere ", 1, -1, 1)
          valor = Replace(valor, "update ", "atualiza ", 1, -1, 1)
          valor = Replace(valor, "table ", "tabela ", 1, -1, 1)
          valor = Replace(valor, "xp_ ", "", 1, -1, 1)
          valor = Replace(valor, "begin ", "iniciar ", 1, -1, 1)
    end if

    SafeVariavelSql = valor

end function


'FUNCAO QUE FILTRA CARACTERES MALICIOSOS'
'-------------------------------------'
function SafeSqlLike(valor)
'-------------------------------------'
	if valor <> "" then
		valor = Replace(valor,"'","")
	''	valor = Replace(valor, "[", "[[]")
		valor = Replace(valor, "%", "[%]")
		valor = Replace(valor, "_", "[_]")
	end if

	SafeSqlLike = valor

end function


'CONVERTE O ARRAY REQUEST EM UM DICIONARIO COM OS DADOS FILTRADOS DE CARACTER MALICIOSO'
'-----------------------------------------------'
function ConverteRequestParaDicionario(requestObj)
'-----------------------------------------------'
	Dim d, item, fieldName, fieldValue
	Set d = Server.CreateObject("Scripting.Dictionary")

	For Each item In requestObj.Form
	    fieldName = item
	    fieldValue = SafeVariavelSql(requestObj.Form(Item))
  
	    d.Add fieldName, fieldValue
	Next 

	For Each item In requestObj.QueryString
	    fieldName = item
	    fieldValue = SafeVariavelSql(requestObj.QueryString(Item))

	    d.Add fieldName, fieldValue
	Next

	set ConverteRequestParaDicionario = d

end function

'FUNCAO QUE FILTRA STRING REMOVENDO CARACTERES QUE NAO SEJAM ALFANUMERICOS'
'-------------------------------------'
function LimpaCampoParaAlfaNumerico(valor)
'-------------------------------------'
	if valor <> "" then
		Set regEx = New RegExp
		regEx.Global = True
		regEx.Pattern = "[^A-Za-z0-9 ]"
		LimpaCampoParaAlfaNumerico = regEx.Replace(valor,"")
	else
		LimpaCampoParaAlfaNumerico = ""
	end if

end function

'FUNCAO QUE FILTRA STRING REMOVENDO CARACTERES QUE NAO SEJAM NUMERICOS'
'-------------------------------------'
function LimpaCampoParaNumerico(valor)
'-------------------------------------'
	if valor <> "" then
		Set regEx = New RegExp
		regEx.Global = True
		regEx.Pattern = "[\D_]"
		LimpaCampoParaNumerico = regEx.Replace(valor,"")
	else
		LimpaCampoParaNumerico = ""
	end if

end function


Function File_Get_Contents(strFile)
	' Remote File
	If Left(strFile, 7) = "http://" Or Left(strFile, 8) = "https://" Then
		Set objXML = Server.CreateObject("Microsoft.XMLHTTP")
		' Use this line if above errors
		'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.Open "GET", strFile, False
		objXML.Send()
		File_Get_Contents = objXML.ResponseText()
		Set objXML = Nothing
	' Local File
	Else
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(strFile, 1)
		File_Get_Contents = objFile.ReadAll()
		Set objFile = Nothing
		Set objFSO = Nothing
	End If
End Function


'FUNCOES DE VALIDACAO DE FORMULARIO

function ValidaCampoObrigatorio(campo)
	if len(trim(campo)) < 1 then
		ValidaCampoObrigatorio = false
		exit function
	end if
	ValidaCampoObrigatorio = true
end function 

function ValidaNumero(campo)
	ValidaNumero = isNumeric(campo)
end function

function ValidaTamanhoString(campo, min, max)
	tamanho = len(trim(campo))
	if tamanho < min then
		ValidaTamanhoString = false
		exit function
	end if
	if max <> "" then
		if tamanho > max then
			ValidaTamanhoString = false
			exit function
		end if
	end if
	ValidaTamanhoString = true	
end function

function ValidaNome(nome) 
        Set regEx = New RegExp 
        regEx.Pattern = "^([a-z\u00C0-\u00ffA-Z]{2,}([a-z\u00C0-\u00ffA-Z]+)*)(\s([a-z\u00C0-\u00ffA-Z]+([a-z\u00C0-\u00ffA-Z]+){0,}))+$" 
        ValidaNome = regEx.Test(trim(nome)) 
end function

function ValidaData(campo, min, max)
	if len(trim(campo)) <> 10 then
		ValidaData = false
		exit function
	end if 

	dmy = split(campo,"/")
	data = dmy(2) & "-" & dmy(1) & "-" & dmy(0)	

	if isDate(data) = false then
		ValidaData = false
		exit function
	end if
	if min <> "" then
		min = cdate(min)
		if data < min then
			ValidaData = false
			exit function
		end if
	end if
	if max <> "" then
		max = cdate(max)
		if data > max then
			ValidaData = false
			exit function
		end if
	end if
	ValidaData = true
end function

function ValidaDataNascimento(campo)
	if len(trim(campo)) <> 10 then
		ValidaDataNascimento = false
	end if 
	
	dmy = split(campo,"/")
	data = cdate(dmy(2) & "-" & dmy(1) & "-" & dmy(0))

	if isDate(data) = false then
		ValidaDataNascimento = false
		exit function
	end if
	if data >= Date then
		ValidaDataNascimento = false
		exit function
	end if
	if DateDiff("yyyy",data, Date) > 120 then
		ValidaDataNascimento = false
		exit function
	end if
	ValidaDataNascimento = true
end function

function ValidaEmail(email) 
        Set regEx = New RegExp 
        regEx.Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w{2,}$" 
        ValidaEmail = regEx.Test(trim(email)) 
end function

function ValidaCep(cep)
    cep = replace(cep,"-","")
    if len(cep) <> 8 then
    	ValidaCep = false
    	exit function
    end if
    Set regEx = New RegExp
    regEx.Pattern = "^[0-9]{5}[0-9]{3}$" 
    ValidaCep = regEx.Test(cep) 
end function

function ValidaCpf(cpf)
    Dim multiplic1, multiplic2
    multiplic1=Array(10, 9, 8, 7, 6, 5, 4, 3, 2)
    multiplic2=Array(11, 10, 9, 8, 7, 6, 5, 4, 3, 2 )
    Dim tempCpf,digit,sum,remainder,i,RegXP
    cpf = Trim(cpf)
    cpf = Replace(cpf,".", "")
    cpf = Replace(cpf,"-", "")
    if (Len(cpf) <> 11) Then
        ValidateCPFNew = false
        exit function
    else
        tempCpf = Left (cpf, 9)
        sum = 0

        Dim intCounter
        Dim intLen 
        Dim arrChars()

        intLen = Len(tempCpf)-1
        redim arrChars(intLen)

        For intCounter = 0 to intLen
            arrChars(intCounter) = Mid(tempCpf, intCounter + 1,1)
        Next

        i=0
        For i = 0 to 8
            sum =sum + CInt(arrChars(i)) * multiplic1(i)
        Next

        remainder = sum Mod 11
        If (remainder < 2) Then
            remainder = 0
        else
            remainder = 11 - remainder
        End If

        digit = CStr(remainder)
        tempCpf = tempCpf & digit
        sum = 0

        intLen = Len(tempCpf)-1
        redim arrChars(intLen)
        intCounter= 0
        For intCounter = 0 to intLen
            arrChars(intCounter) = Mid(tempCpf, intCounter + 1,1)
        Next
        i=0
        For i = 0 to 9
            sum =sum + CInt(arrChars(i)) * multiplic2(i)
        Next        
        remainder = sum Mod 11

        If (remainder < 2) Then
            remainder = 0
        else
            remainder = 11 - remainder
        End If      
        digit = digit & CStr(remainder)

        Set RegXP=New RegExp
            RegXP.IgnoreCase=1
            RegXP.Pattern=digit & "$"

        If RegXP.test(cpf) Then 
            ValidaCpf = true
        else
            ValidaCpf = false
        end if
    end if
end Function


function AbreviarNome(pNome, ptam)

    pNome = Trim(pNome)
    If Len(pNome) <= ptam Then
        AbreviarNome = pNome
        Exit Function
    End If
    '
    'Retira os pronomes
    '
    pNome = Replace(pNome, " DOS ", " ")
    pNome = Replace(pNome, " DO ", " ")
    pNome = Replace(pNome, " DAS ", " ")
    pNome = Replace(pNome, " DA ", " ")
    pNome = Replace(pNome, " DE ", " ")
    pNome = Replace(pNome, " E ", " ")
    If Len(pNome) <= ptam Then
        AbreviarNome = pNome
        Exit Function
    End If
    '
    'Determina o número de "nomes" do nome
    '
    Espaco1 = InStr(pNome, " ")
    Espaco3 = InStrRev(pNome, " ")
    if Espaco1 = 0 Then
    	Espaco2 = 0
    Else
    	Espaco2 = InStrRev(pNome, " ", Espaco3 - 1)
    end if
    '
    'Um nome
    '
    If Espaco1 = 0 Then
        AbreviarNome = Left(pNome, ptam)
        Exit Function
    End If
    '
    'Dois nomes
    '
    If Espaco1 = Espaco2 And Espaco1 = Espaco3 Then
        Primeiro = Left(pNome, 1) + " "
        Final = Trim(Right(pNome, Len(pNome) - Espaco1 + 1))
        AbreviarNome = Left(Primeiro + Final, ptam)
        Exit Function
    End If
    '
    'Três nomes
    '
    'If Espaco1 = Espaco2 Then
    '   Primeiro = Left(pNome, 1) + ". "
    '   Final = Trim(Right(pNome, Len(pNome) - Espaco1 + 1))
    '   AbreviarNome = Left(Primeiro + Final, ptam)
    '   Exit Function
    'End If
    If Espaco1 = Espaco2 Then
        Primeiro = Trim(Left(pNome, Espaco1 - 1)) + " "
        Final = Trim(Right(pNome, Len(pNome) - Espaco3 + 1))
        Meio = Trim(Mid(pNome, Espaco1 + 1, Espaco3 - 1))
        Meio = Left(Meio, 1) + " "
        If Len(Primeiro + Meio + Final) > ptam Then
            Primeiro = Left(Primeiro, 1) + " "
        End If
        AbreviarNome = Left(Primeiro + Meio + Final, ptam)
        Exit Function
    End If
    '
    'Mais de três
    '
    Primeiro = Left(pNome, Espaco1 - 1)
    Final = Trim(Right(pNome, Len(pNome) - Espaco3 + 1)) '2

    if (Espaco3 - Espaco1 - 1) > 0 then
         Meio = Mid(pNome, Espaco1 + 1, Espaco3 - Espaco1 - 1) '2
    else
         Meio = ""
    end if
    
    conta = 0
    NovoMeio = ""
    Do
        Espaco2 = InStr(conta + 1, Meio, " ")
        If Espaco2 = 0 Then
            Espaco2 = Len(Meio)
        End If
        NomeMeio = Trim(Mid(Meio, conta + 1, Espaco2 - conta))
        'If UCase(NomeMeio) = "DOS" Or _
            '   UCase(NomeMeio) = "DO" Or _
            '   UCase(NomeMeio) = "DAS" Or _
            '   UCase(NomeMeio) = "DA" Or _
            '   UCase(NomeMeio) = "DE" Or _
            '   UCase(NomeMeio) = "E" Then
        '   NomeMeio = ""
        ' Else
        NomeMeio = Left(NomeMeio, 1) + " "
        'End If
        conta = Espaco2
        NovoMeio = NovoMeio + NomeMeio
    Loop Until conta = Len(Meio)
    
    If Len(Primeiro + " " + NovoMeio + Final) > ptam Then
        Primeiro = Left(Primeiro, 1)
    End If
    '
    'Is this the end
    '
    AbreviarNome = Left(Primeiro + " " + NovoMeio + Final, ptam)

End Function






%>