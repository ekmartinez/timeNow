Attribute VB_Name = "Module1"

Sub Time()

    Dim Time As Date
    Time = Now()
    Range("a1") = Time

End Sub

'=IF(AND(COUNT(B2:D2)>0,COUNT(A2)<1),"Entre Numero de Departamento", "")