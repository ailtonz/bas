Attribute VB_Name = "M�dulo1"
Sub fibo()

    penul = Day(Time)
    ultim = Month(Time)
    
    ini = Year(Date) \ 100
    fim = Year(Date) Mod 100
    
    penul = 16
    ultim = 12
    
    ini = 1982 \ 100
    fim = 1982 Mod 100
    
    
    i = 3
    
    While Len(senha) < 8
        senha = ((penul * ini) + ultim) - fim
        penul = ultim
        ultim = senha
        i = i + 1
    Wend
    
    MsgBox senha

End Sub
