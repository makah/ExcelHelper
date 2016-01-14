VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OCFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Representa uma nova planilha
'
' @author Mauricio Arieira
' @date Julho 2013
'
' @version 1.2


' Caminho do Excel que será aberto.
Private path As String

' Variável que define se o arquivo foi aberto pelo objeto ou se já estava aberto.
Private openedByMe As Boolean

' WorkBook foi requisitado a abertura do novo arquivo Excel.
Public oldWorkbook As Workbook

' WorkBook do novo arquivo Excel.
Public newWorkbook As Workbook

' Guarda se o arquivo foi aberto como leitura.
Private pIsReadOnly As Boolean

' Abre uma planilha que se encontra em outro arquivo.
' Abre caso a planilha não esteja aberta. Caso contrário ativa a planilha solicitada.
' @param In filePath: O caminho da planilha
' @param Optional In isReadOnly: Define se o arquivo será aberto somente para leitura.
Sub OpenNewFile(ByVal FilePath As String, Optional ByVal isReadOnly As Boolean = True)
    Dim fileName As String
    Dim actualScreenUpdate As Boolean, actualDisplayAlert As Boolean

    actualScreenUpdate = Application.ScreenUpdating
    actualDisplayAlert = Application.DisplayAlerts
    Application.ScreenUpdating = False
    If IsNull(oldWorkbook) Or IsNull(newWorkbook) Then
        MsgBox "Classe está sendo utilizada para referenciar outra planilha. Feche o arquivo antes de abrir outro."
    End If
    
    Set oldWorkbook = ActiveWorkbook
    pIsReadOnly = isReadOnly
    
    'Se já estiver aberto e for readOnly não precisa abrir.
    fileName = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
    
    On Error GoTo ElseLabel:
    Windows(fileName).Activate
    Set newWorkbook = ActiveWorkbook
    openedByMe = False
    GoTo EndLabel
    
    On Error GoTo 0

    'Senão abre o arquivo
ElseLabel:
    If Dir(FilePath) = "" Then
        err.Raise vbObjectError + 101, "OCFile", "File not found: '" & FilePath & "'"
    End If
    
    Workbooks.Open fileName:="" & FilePath & "", UpdateLinks:=False, ReadOnly:=isReadOnly
    Application.DisplayAlerts = actualDisplayAlert
    openedByMe = True
        
EndLabel:
    Set newWorkbook = ActiveWorkbook
    oldWorkbook.Activate
    Application.ScreenUpdating = actualScreenUpdate
End Sub

' Fecha o arquivo aberto. Caso o objeto não tenha sido responsável por abri-la
'então ele não fecha.
Sub CloseFile(Optional ByVal saveChanges As Boolean = True)
    On Error Resume Next
    If openedByMe Then
        oldWorkbook.Activate
        If pIsReadOnly Then
            Call newWorkbook.Close(False)
        Else
            newWorkbook.Close (saveChanges)
        End If
        oldWorkbook.Activate
    End If
    
    Set oldWorkbook = Nothing
    Set newWorkbook = Nothing
    On Error GoTo 0
End Sub
