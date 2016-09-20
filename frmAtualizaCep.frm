VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAtualizaCep 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localiza Cep - Indicar Atualização"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   15000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnVoltar 
      Height          =   495
      Left            =   14280
      Picture         =   "frmAtualizaCep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Voltar"
      Top             =   7440
      Width           =   495
   End
   Begin SHDocVwCtl.WebBrowser wbSiteAtualizaCep 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      ExtentX         =   26485
      ExtentY         =   12726
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmAtualizaCep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnVoltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo trata_erro
    
    wbSiteAtualizaCep.Navigate Trim("https://viacep.com.br/cep/")
    
    Exit Sub
    
trata_erro:
       MsgBox Err.Description
End Sub
