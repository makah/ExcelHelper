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

' @author Mauricio Arieira
' @date Julho 2013
'
' @version 1.1

' new Excel workbook
Private Path As String

Public oldWorkbook As Workbook
Public newWorkbook As Workbook

Private pIsReadOnly As Boolean
Private openedByMe As Boolean

' Abre uma planilha que se encontra em outro arquivo.
' Abre caso a planilha não esteja aberta. Caso contrário ativa a planilha solicitada.
' @param In filePath: O caminho da planilha
' @param Optional In isReadOnly: Define se o arquivo será aberto somente para leitura.
Sub Open(ByVal filePath As String, Optional ByVal isReadOnly As Boolean = True)
    Dim fileName As String
    Dim actualScreenUpdate As Boolean

    actualScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If IsNull(oldWorkbook) Or IsNull(newWorkbook) Then
        MsgBox "Already in use."
    End If
    
    Set oldWorkbook = ActiveWorkbook
    pIsReadOnly = isReadOnly
    
    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    
    On Error GoTo ElseLabel:
    Windows(fileName).Activate
    Set newWorkbook = ActiveWorkbook
    openedByMe = False
    GoTo EndLabel
    
    On Error GoTo 0

ElseLabel:
    Workbooks.Open fileName:="" & filePath & "", UpdateLinks:=False, ReadOnly:=isReadOnly
    openedByMe = True
        
EndLabel:
    Set newWorkbook = ActiveWorkbook
    oldWorkbook.Activate
    Application.ScreenUpdating = actualScreenUpdate
End Sub

Sub Close(Optional ByVal saveChanges As Boolean = True)
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
