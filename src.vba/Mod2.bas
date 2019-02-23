Function Ano_2_digitos(ano)
    str_ano = CStr(ano)
    str_ano = Right(str_ano, 2)
    response = CInt(str_ano)
    Ano_2_digitos = CStr(response)
End Function

Function T_1(ano)
    response = "1T" & Ano_2_digitos(ano)
    T_1 = response
End Function

Function T_2(ano)
    response = "2T" & Ano_2_digitos(ano)
    T_2 = response
End Function

Function T_3(ano)
    response = "3T" & Ano_2_digitos(ano)
    T_3 = response
End Function

Function T_4(ano)
    response = "4T" & Ano_2_digitos(ano)
    T_4 = response
End Function

Function YTD(ano)
    response = "YTD" & Ano_2_digitos(ano)
    YTD = response
End Function
