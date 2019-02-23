Sub Tela_Cheia()

Application.DisplayFullScreen = True

End Sub

Sub Tela_Normal()

Application.DisplayFullScreen = False

End Sub

Sub Operador_1()

Application.ScreenUpdating = False

Rows("9:45").Select
Selection.EntireRow.Hidden = False

Rows("46:99").Select
Selection.EntireRow.Hidden = True

ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = True
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = False

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Corretor_1()

Application.ScreenUpdating = False

Rows("9:45").Select
Selection.EntireRow.Hidden = True

Rows("46:99").Select
Selection.EntireRow.Hidden = False
   
ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = False
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = True

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Operador_2()

Application.ScreenUpdating = False

Rows("11:82").Select
Selection.EntireRow.Hidden = False

Rows("83:189").Select
Selection.EntireRow.Hidden = True

ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = True
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = False

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Corretor_2()

Application.ScreenUpdating = False

Rows("11:82").Select
Selection.EntireRow.Hidden = True

Rows("83:189").Select
Selection.EntireRow.Hidden = False

ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = False
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = True
    
Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Operador_3()

Application.ScreenUpdating = False

Rows("10:63").Select
Selection.EntireRow.Hidden = False

Rows("64:143").Select
Selection.EntireRow.Hidden = True

ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = True
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = False

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Corretor_3()

Application.ScreenUpdating = False

Rows("10:63").Select
Selection.EntireRow.Hidden = True

Rows("64:143").Select
Selection.EntireRow.Hidden = False
 
ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = False
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = True
    
Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Operador_Total()

Application.ScreenUpdating = False

Rows("16:43").Select
Selection.EntireRow.Hidden = False

Rows("44:79").Select
Selection.EntireRow.Hidden = True

ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = True
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = False

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub Corretor_Total()

Application.ScreenUpdating = False

Rows("16:43").Select
Selection.EntireRow.Hidden = True

Rows("44:79").Select
Selection.EntireRow.Hidden = False
    
ActiveSheet.Shapes.Range(Array("VISAO_OPERADOR")).Visible = False
ActiveSheet.Shapes.Range(Array("VISAO_CORRETOR")).Visible = True
    
Range("A1").Select

Application.ScreenUpdating = True

End Sub
Sub Prev_HideandShow()

'Application.ScreenUpdating = False

Rows("42:45").Select

If Selection.EntireRow.Hidden = True Then
Selection.EntireRow.Hidden = False
Else
Selection.EntireRow.Hidden = True
End If

Rows("41:41").Select
Selection.EntireRow.Hidden = False
Rows("47:49").Select
Selection.EntireRow.Hidden = False

End Sub

Sub Ajusta_Rateio_Rede_Receita()
'
' Ajusta_Rateio_Rede_Receita Macro
'

'

Application.ScreenUpdating = False

    Range("D8").Select
    ActiveCell.FormulaR1C1 = "='Despesas Administrativas'!R[9]C"
    Range("D8").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.End(xlToLeft).Select
    Range("D13").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.FormulaR1C1 = "0%"
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "=100%-SUM(R[-8]C:R[-1]C)"
    Range("D16").Select
    Selection.Copy
    Range("E16").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.End(xlToLeft).Select
    Range("E16").Select
    Application.CutCopyMode = False
    Range("D16").Select
    Selection.End(xlUp).Select
End Sub
Sub Habitacional1()

Sheets("Bloco I - Result.").Select
Call Operador_1
Range("C8").Select

End Sub
Sub Consorcio1()

Sheets("Bloco I - Result.").Select
Call Operador_1
Range("C36").Select

End Sub
Sub Auto1()

Sheets("Bloco II - Result.").Select
Call Operador_2
Range("C11").Select

End Sub
Sub Residencial1()

Sheets("Bloco II - Result.").Select
Call Operador_2
Range("C37").Select

End Sub
Sub Empresarial1()

Sheets("Bloco II - Result.").Select
Call Operador_2
Range("C55").Select

End Sub
Sub Rural1()

Sheets("Bloco II - Result.").Select
Call Operador_2
Range("C73").Select

End Sub
Sub Vida1()

Sheets("Bloco III - Result.").Select
Call Operador_3
Range("C9").Select

End Sub
Sub Prestamista1()

Sheets("Bloco III - Result.").Select
Call Operador_3
Range("C36").Select

End Sub
Sub Previdencia1()

Sheets("Bloco III - Result.").Select
Call Operador_3
Range("C54").Select

End Sub
Sub Habitacional2()

Sheets("Bloco I - Result.").Select
Call Corretor_1
Range("C46").Select

End Sub
Sub Consorcio2()

Sheets("Bloco I - Result.").Select
Call Corretor_1
Range("C84").Select

End Sub
Sub Auto2()

Sheets("Bloco II - Result.").Select
Call Corretor_2
Range("C93").Select

End Sub
Sub Residencial2()

Sheets("Bloco II - Result.").Select
Call Corretor_2
Range("C120").Select

End Sub
Sub Empresarial2()

Sheets("Bloco II - Result.").Select
Call Corretor_2
Range("C147").Select

End Sub
Sub Rural2()

Sheets("Bloco II - Result.").Select
Call Corretor_2
Range("C174").Select

End Sub
Sub Vida2()

Sheets("Bloco III - Result.").Select
Call Corretor_3
Range("C74").Select

End Sub
Sub Prestamista2()

Sheets("Bloco III - Result.").Select
Call Corretor_3
Range("C101").Select

End Sub
Sub Previdencia2()

Sheets("Bloco III - Result.").Select
Call Corretor_3
Range("C128").Select

End Sub