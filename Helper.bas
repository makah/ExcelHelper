Attribute VB_Name = "Helper"
Option Explicit

' @author Mauricio Arieira
' @date Janeiro 2016
'
' @version 2.0

' Ordena a tabela da Range selecionada. Lembrando que a tabela não estar com as colunas no 'AutoFilter'.
' @param In currentSheet: Sheet onde a tabela se encontra.
' @param In tableRange: A tabela.
' @param In columnIndex: O índice da coluna da tabela. Caso queira ordenar a primeira coluna 'columnIndex' = 1
' @param In ascending: True para ordenação acendente, False para ordenação descendente
Sub Ordenar(ByVal currentSheet As Worksheet, ByVal tableRange As Range, ByVal columnIndex As Integer, ascending As Boolean)
    Dim orderBy As Integer
    orderBy = IIf(ascending, xlAscending, xlDescending)
    
    tableRange.AutoFilter
    With currentSheet
        .AutoFilter.Sort.SortFields.Clear
        .AutoFilter.Sort.SortFields.Add _
            Key:=Cells(tableRange.Row, tableRange.column + (columnIndex - 1)), _
            SortOn:=xlSortOnValues, order:=orderBy, DataOption:=xlSortNormal
        With .AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    tableRange.AutoFilter
End Sub


' Importa o arquivo para a celula destino utilizando tab para separação de colunas.
' @param In filePath: O caminho onde o arquivo está localizado.
' @param In columnDataType: Um array com os tipos de cada coluna. Exemplo: Array(xlTextFormat, xlSkipColumn, xlGeneralFormat)
' @param In destinationCell: A celula de destino da importação.
' @param In isDotDecimnalSeparator: True caso o aquivo utilize '.' como separador decimal e falso caso utilize ','.
Sub ImportFile(ByVal filePath As String, ByVal columnDataType As Variant, _
        ByVal destinationCell As Range, ByVal isDotDecimnalSeparator As Boolean)
    Dim qt As QueryTable
    Dim destinationWorkSheet As Worksheet

    Set destinationWorkSheet = destinationCell.Parent

    With destinationWorkSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filePath, Destination:=destinationCell)
        .name = filePath
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = columnDataType
        .TextFileDecimalSeparator = IIf(isDotDecimnalSeparator, ".", ",")
        .TextFileThousandsSeparator = IIf(isDotDecimnalSeparator, ",", ".")
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    For Each qt In destinationWorkSheet.QueryTables
        qt.Delete
    Next qt
End Sub


' Converte o caminho do arquivo para uma string que pode ser adicionada na fórmula
' @param In path: O caminho do arquivo
' @return A string modificada
' @remarks A função não verifica se o arquivo existe.
Function PathToFormula(ByVal path As String, ByVal sheetName As String) As String
    Dim arr As Variant, lenght As Integer, filePath As String, directoryPath As String

    arr = Split(path, "\")
    lenght = UBound(arr)
    filePath = arr(lenght)
    arr(lenght) = ""
    PathToFormula = StringFormat("'{0}[{1}]{2}'!", Join(arr, "\"), filePath, sheetName)
End Function

' Gera uma string com usando o formato do .NET - {0}, {1}, {2} ...
' @param In strValue as String: A string utilizando o padrão {n} onde n é o enésimo parâmetro da variável 'arrParames'
' @param In arrParames as Variant: Um vetor com os argumentos que serão substituidos na string 'strValue'
' @return as String: A string já formatada.
' @example:  call StringFormat("My name is {0} {1}. Hey!", "Mauricio", "Arieira")
Public Function StringFormat(ByVal strValue As String, ParamArray arrParames() As Variant) As String
    Dim i As Integer

    For i = LBound(arrParames()) To UBound(arrParames())
        strValue = Replace(strValue, "{" & CStr(i) & "}", CStr(arrParames(i)))
    Next
    
    StringFormat = strValue
End Function

' Encontra todas as ocorrências de um valor em um intervalo
' @param In value as String: O valor a ser procurado no intervalo
' @param In theRange as Range: O intervalo
' @return as Double(): O vetor contendo as linhas que foram encontradas.
Public Function MatchAll(ByVal value As String, ByVal theRange As Range) As Double()
    Dim index As Long, rFoundCell As Range, total As Integer, results() As Double
         
    total = WorksheetFunction.CountIf(theRange, value)
    If total = 0 Then
        Exit Sub
    End If
    ReDim results(total)
    
    Set rFoundCell = theRange.Cells(1, 1)
    For index = 1 To total
         
        Set rFoundCell = theRange.Find(What:=value, After:=rFoundCell, _
                LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False)
        
        results(index) = rFoundCell.Row
    Next index
    
    MatchAll = results
End Function
