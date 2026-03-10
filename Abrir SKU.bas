Attribute VB_Name = "Módulo1"
Sub GerarCodVar()
Attribute GerarCodVar.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False

    '''''''''''''''''''''''' DECLARAÇÕES ''''''''''''''''''''''''''
    Dim dd As Worksheet
    Dim dg As Worksheet
    Set dd = Sheets("Dataset_Dimensoes")
    Set dg = Sheets("Geral")
    
    ''''''''''''''''''''''''''' LIMPEZA '''''''''''''''''''''''''''
        ' Aba Geral
        dg.Select
            last_row_clear = dg.[D1].End(xlDown).Row
            dg.Range("D2:G" & last_row_clear).Clear
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''' VALIDAÇÃO CÓDIGO TER 6 DIGITOS '''''''''''''''
        last_row = dg.[A1].End(xlDown).Row
        For i = 2 To last_row
            If Len(Range("A" & i)) <> 6 Then
                MsgBox ("Código " & cod & " na linha " & i & " não tem 6 digitos. Favor corrigir e executar novamente!")
            End If
        Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    '''''''''''' LOOPING PARA GERAR CÓDIGOS VARIANTES '''''''''''''
    For i = 2 To last_row
        cod = Range("A" & i)
        num_var = WorksheetFunction.CountIf(dd.Range("B:B"), cod)
        preco = Range("B" & i)
        If num_var = 0 Then
            setor = ""
        Else
            setor = Application.WorksheetFunction.VLookup(dg.Range("A" & i).Text, dd.Range("B:C"), 2, False)
        End If
        If num_var < 2 Then
            last_row_out = dg.[D1048576].End(xlUp).Row + 1
                Range("D" & last_row_out) = cod
                Range("E" & last_row_out) = cod
                Range("F" & last_row_out) = preco
                Range("G" & last_row_out) = setor
        Else
            last_row_out = dg.[D1048576].End(xlUp).Row + 1
                    dg.Range("D" & last_row_out) = cod
                    dg.Range("E" & last_row_out) = cod
                    dg.Range("F" & last_row_out) = preco
                    dg.Range("G" & last_row_out) = setor
                For j = 1 To (num_var - 1)
                    tamanho = Format(j, "000")
                    cod_tam = cod & tamanho
                    dg.Range("D" & last_row_out + j) = cod
                    dg.Range("E" & last_row_out + j) = cod_tam
                    dg.Range("F" & last_row_out + j) = preco
                    dg.Range("G" & last_row_out + j) = setor
                Next
        End If
    Next
    
    ''''''''''''''''''''''''' FORMATAÇÃO ''''''''''''''''''''''''''
    last_row_forms = dg.[D1].End(xlDown).Row
    dg.Range("D1:G" & last_row_forms).HorizontalAlignment = xlCenter
    dg.Range("D1:G" & last_row_forms).BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=vbRed
    dg.Range("D1:G1").Font.Bold = True
    dg.Range("D1:G1").Interior.ColorIndex = 6
    Columns("D:G").EntireColumn.AutoFit
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Application.ScreenUpdating = True

MsgBox ("Pronto!")

End Sub
