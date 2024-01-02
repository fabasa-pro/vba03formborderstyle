VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

' ==========================================================================
' NOTA: para editar o código-fonte, executar o arquivo com a tecla <Shift>
' pressionada para ignorar todo o VBA e entre no aplicativo Microsoft Word.
' ==========================================================================

Option Explicit

Private Declare PtrSafe Function FindWindow Lib "User32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWdd As LongPtr, ByVal nIndex As GWL, ByVal dwNewLong As Long) As Long

Private Enum GWL
    GWL_EXSTYLE = -20     ' Define um novo estilo de janela estendida .
    GWL_HINSTANCE = -6    ' Define um novo identificador de instância do aplicativo.
    GWL_ID = -12          ' Define um novo identificador da janela filho.
    GWL_STYLE = -16       ' Define um novo estilo de janela .
    GWL_USERDATA = -21    ' Define os dados do usuário associados à janela.
    GWL_WNDPROC = -4      ' Define um novo endereço para o procedimento da janela.
End Enum

Private Sub UserForm_Initialize()

    Dim hWnd As LongPtr
    hWnd = FindWindow(vbNullString, Me.Caption)
        
    Dim FormBorderStyle As Integer
    FormBorderStyle = 4
    
    Select Case FormBorderStyle
        Case 0                                          ' 0-None
            SetWindowLong hWnd, GWL_EXSTYLE, &H50000
            SetWindowLong hWnd, GWL_STYLE, &H6010000
            Me.BackColor = RGB(255, 255, 255)
        Case 1                                          ' 1-FixedSingle
            SetWindowLong hWnd, GWL_EXSTYLE, &H50100
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 2                                          ' 2-Fixed3D
            SetWindowLong hWnd, GWL_EXSTYLE, &H50300
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 3                                          ' 3-FixedDialog
            SetWindowLong hWnd, GWL_EXSTYLE, &H50101
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 4                                          ' 4-Sizable
            SetWindowLong hWnd, GWL_EXSTYLE, &H50100
            SetWindowLong hWnd, GWL_STYLE, &H6CF0000
        Case 5                                          ' 5-FixedToolWindow
            SetWindowLong hWnd, GWL_EXSTYLE, &H50180
            SetWindowLong hWnd, GWL_STYLE, &H6CB0000
        Case 6                                          ' 6-SizableToolWindow
            SetWindowLong hWnd, GWL_EXSTYLE, &H50180
            SetWindowLong hWnd, GWL_STYLE, &H6CF0000
    End Select
    
End Sub

Private Sub UserForm_Terminate()

    Project.ThisDocument.Application.Visible = True                                                    ' Ocultar ou mostrar aplicativos.
    Project.ThisDocument.Application.Quit SaveChanges:=wdSaveChanges, OriginalFormat:=wdWordDocument   ' Salvar e fechar tudo.
    
End Sub
