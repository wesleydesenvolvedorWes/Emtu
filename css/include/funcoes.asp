<%
    Function removeAcentos(Palavra) 
        cacento = "אבגדהטיךכלםמןעףפץצשתûüְֱֲֳִָֹÊֻּֽ־ׂ׃װױײÙÚÛÜחַסׁ^~÷×´`'" 
        sacento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN" 
        texto = "" 
        If Palavra <> "" Then 
            For x = 1 To Len(Palavra) 
                letra = Mid(Palavra, x, 1) 
                pos_acento = InStr(cacento, letra) 
                If pos_acento > 0 Then 
                    letra = Mid(sacento, pos_acento, 1) 
                End If 
                texto = texto & letra 
            Next 
            TiraAcento = texto 
        End If 
    End Function
%>