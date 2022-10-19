Attribute VB_Name = "Módulo1"

Sub geral()

Application.ScreenUpdating = False

Dim msgb As String

msgb = MsgBox("Processar atualização de dados?", vbYesNo, "VALIDAÇÃO DE ATIVAÇÃO DE MACROS")

If msgb = 6 Then

    Call bd_fat_m0
    Call bd_fat_m0_2
    Call graficos_de_envio

Else
End If

    Sheets("PAINEL DE ATUALIZAÇÃO").Select
    Range("G4").Select

Application.ScreenUpdating = True

End Sub


' ==========================================================================================
' ==========================================================================================
' ==========================================================================================

Sub bd_fat_m0()

Application.ScreenUpdating = False

' Variáveis
Dim connection As WorkbookConnection
Dim query As WorkbookQuery

    Sheets("BD - FAT M0").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B5").Select

' CARREGAMENTO DO DS

    Sheets("BD - FAT M0").Select
    Path = ActiveWorkbook.Path & "\FAT M0 ATÉ D-1 - CLIENTE.EQUIPE.csv"

    Application.CutCopyMode = False
    ActiveWorkbook.Queries.Add Name:="FAT M0 ATÉ D-1 - CLIENTE EQUIPE", Formula _
        := _
        "let" & Chr(13) & "" & Chr(10) & "    Fonte = Csv.Document(File.Contents(""" & Path & """),[Delimiter="","", Columns=33, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Cabeçalhos Promovidos"" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Tipo Alterado"" = Table.TransformColumnTypes(#""Cabeçalhos Promo" & _
        "vidos"",{{""%Part. Acum."", type number}, {""Código Cliente"", Int64.Type}, {""Cliente : Equipe"", type text}, {""Quantidade Pedida"", Int64.Type}, {""Quantidade Atendida"", Int64.Type}, {""Quantidade Cortada"", Int64.Type}, {""% Qtd Corte"", type number}, {""Preç.Méd. Tabela"", type number}, {""Preç.Méd. Pedido"", type number}, {""% Méd. Descto"", type number}, {""" & _
        "Total Pedido"", type number}, {""Promoções Total Atendido"", type number}, {""Promoções %Prom/Vda."", type number}, {""Total Cortado"", type number}, {""%  Corte"", type number}, {""Total Atendido"", type number}, {""%Partic. Total"", type number}, {""Lucratividade Total"", type number}, {""Margem Cadastro"", type text}, {""Lucrat.% Margem"", type number}, {""Prazo " & _
        "Médio"", type number}, {""Contag. Pedidos"", Int64.Type}, {""Contag. Prods."", Int64.Type}, {""Contag. Clientes"", Int64.Type}, {""Vlr Flex Positivo"", type number}, {""Vlr Flex Negativo"", type number}, {""Vlr Flex Saldo"", type number}, {""Peso Bruto"", type number}, {""Peso Brt %Partic."", type number}, {""Peso Líquido"", type number}, {""Custo Liquido"", type nu" & _
        "mber}, {""Comissão"", type number}, {""Comissão Televenda"", Int64.Type}})," & Chr(13) & "" & Chr(10) & "    #""Dividir Coluna por Delimitador"" = Table.SplitColumn(#""Tipo Alterado"", ""Cliente : Equipe"", Splitter.SplitTextByDelimiter("":"", QuoteStyle.Csv), {""Cliente : Equipe.1"", ""Cliente : Equipe.2""})," & Chr(13) & "" & Chr(10) & "    #""Tipo Alterado1"" = Table.TransformColumnTypes(#""Dividir Coluna por Delimita" & _
        "dor"",{{""Cliente : Equipe.1"", type text}, {""Cliente : Equipe.2"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Colunas Renomeadas"" = Table.RenameColumns(#""Tipo Alterado1"",{{""Cliente : Equipe.1"", ""Cliente""}, {""Cliente : Equipe.2"", ""Equipe""}})," & Chr(13) & "" & Chr(10) & "    #""Texto Aparado"" = Table.TransformColumns(#""Colunas Renomeadas"",{{""Cliente"", Text.Trim, type text}, {""Equipe"", Text.Trim," & _
        " type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Texto Aparado"""
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""FAT M0 ATÉ D-1 - CLIENTE EQUIPE"";Extended Properties=""""" _
        , Destination:=Range("$B$5")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [FAT M0 ATÉ D-1 - CLIENTE EQUIPE]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "FAT_M0_ATÉ_D_1___CLIENTE_EQUIPE_2"
        .Refresh BackgroundQuery:=False
    End With
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    Range("B6").Select

On Error Resume Next
For Each connection In ThisWorkbook.Connections
connection.Delete
Next
    
On Error Resume Next
For Each query In ThisWorkbook.Queries
query.Delete
Next

Application.ScreenUpdating = True

End Sub

' ==========================================================================================
' ==========================================================================================
' ==========================================================================================

Sub bd_fat_m0_2()

Application.ScreenUpdating = False

    Sheets("BD - FAT M0 (2)").Select
    Rows("11:150000").Select
    Selection.Delete Shift:=xlUp
    Range("B4").Select
    Sheets("BD - FAT M0").Select
    Range("B6:AI6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BD - FAT M0 (2)").Select
    Range("B4").Select
    ActiveSheet.Paste
    Range("B4").Select
    Application.CutCopyMode = False
    ActiveWorkbook.RefreshAll

Application.ScreenUpdating = True

End Sub

' ==========================================================================================
' ==========================================================================================
' ==========================================================================================

Sub graficos_de_envio()

Application.ScreenUpdating = False

    Sheets("GRÁFICOS DE ENVIO").Select
    Range("N4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("U4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("Y4").Select
    Range("U3:Z3").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("Y3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    Range("U5:Z5,U7:Z7,U9:Z9,U11:Z11,U13:Z13").Select
    Range("U13").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("U4").Select
    Range("N18").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("U18").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("U17:Z17").Select
    Selection.AutoFilter
    Range("Y18").Select
    ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("Y17"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("U17:Z17").Select
    Selection.AutoFilter
    Range("U19:Z19,U21:Z21,U23:Z23,U25:Z25,U27:Z27").Select
    Range("U27").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("U18").Select
    Range("N33").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("AD33").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("AD32:AJ32").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("AF32"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("GRÁFICOS DE ENVIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AF33").Select
    Selection.AutoFilter
    Range("AD33:AJ52").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range( _
        "AD34:AJ34,AD36:AJ36,AD38:AJ38,AD40:AJ40,AD42:AJ42,AD44:AJ44,AD46:AJ46,AD48:AJ48,AD50:AJ50,AD52:AJ52" _
        ).Select
    Range("AD52").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Range("AD33").Select
    Range("A1").Select

Application.ScreenUpdating = True

End Sub

' ==========================================================================================
' ==========================================================================================
' ==========================================================================================

Sub arquivo_de_envio()

Application.ScreenUpdating = False

'Variáveis
Dim msgb As String

msgb = MsgBox("Gerar arquivo de envio?", vbYesNo, "VALIDAÇÃO DE ATIVAÇÃO DE MACROS")

If msgb = 6 Then

    Sheets("PAINEL DE ATUALIZAÇÃO").Select
    ActiveWorkbook.Save
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("PAINEL DE ATUALIZAÇÃO").Range("I14").Value & " - Gestão de Top Compras MS - Dados até dia " & Worksheets("PAINEL DE ATUALIZAÇÃO").Range("J14").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
 
    Sheets("GESTÃO - TOP 300 CLIENTES").Select
    ActiveWindow.DisplayHeadings = True
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    Range("H6:S6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B6").Select
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ActiveWindow.DisplayHeadings = False
    Sheets("GESTÃO - TOP 20 REDES").Select
    Range("E6:P6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B6").Select
    Sheets("GESTÃO - CLIENTES TOP 20 REDES").Select
    ActiveWindow.DisplayHeadings = True
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    Range("I6:T6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B6").Select
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ActiveWindow.DisplayHeadings = False
    Sheets("GESTÃO - TOP 20 REDES").Select
    Range("E1:J1").Select
    Selection.ClearContents
    Range("B6").Select
    Sheets("GESTÃO - TOP 300 CLIENTES").Select
    ActiveWindow.ScrollColumn = 9
    Range("B6").Select
    Sheets("GRÁFICOS DE ENVIO").Select
    Sheets(Array("GRÁFICOS DE ENVIO", "CARTEIRA M0", "BD - RCA X CLI", _
        "BD - RCA X ZNV", "TD - RCAXCLI", "2021", "2022", "FAT TT", "TD - FAT TT - CLI", _
        "TD - FAT TT - RDE", "FAT TT (2)")).Select
    Sheets("GRÁFICOS DE ENVIO").Activate
    Sheets(Array("PAINEL DE ATUALIZAÇÃO", "BD - FAT M0", "BD - FAT M0 (2)", _
        "TD - FAT M0", "GRÁFICOS DE ENVIO", "CARTEIRA M0", "BD - RCA X CLI", _
        "BD - RCA X ZNV", "TD - RCAXCLI", "2021", "2022", "FAT TT", "TD - FAT TT - CLI", _
        "TD - FAT TT - RDE", "FAT TT (2)", "TD - CARTEIRA M0")).Select
    Sheets("PAINEL DE ATUALIZAÇÃO").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("GESTÃO - TOP 300 CLIENTES").Select
    Range("B6").Select
    ActiveWorkbook.Save

Else
End If

Application.ScreenUpdating = True

End Sub