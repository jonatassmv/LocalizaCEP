VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisa 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localiza CEP"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCriteriosPesq 
      BackColor       =   &H8000000B&
      Caption         =   "Filtros de Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   14535
      Begin VB.CommandButton btnAtualizaCep 
         Height          =   495
         Left            =   13680
         Picture         =   "frmPesquisa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Indique um cep que necessita ser atualizado"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton btnPesquisar 
         Height          =   495
         Left            =   9240
         Picture         =   "frmPesquisa.frx":04FB
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   $"frmPesquisa.frx":0933
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtRua 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4200
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox cboUF 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid fxgDadosRuas 
         Height          =   6135
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   10821
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Label lblInstrucaoDbClick 
         BackColor       =   &H8000000B&
         Caption         =   $"frmPesquisa.frx":09C9
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   7560
         Visible         =   0   'False
         Width           =   14295
      End
      Begin VB.Label Label3 
         Caption         =   "Rua"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   795
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Cidade"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "UF"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7440
         TabIndex        =   6
         Top             =   795
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    CarregaUF
End Sub

Private Sub CarregaUF()

    cboUF.AddItem ("")
    cboUF.AddItem ("AC")
    cboUF.AddItem ("AL")
    cboUF.AddItem ("AP")
    cboUF.AddItem ("AM")
    cboUF.AddItem ("BA")
    cboUF.AddItem ("CE")
    cboUF.AddItem ("DF")
    cboUF.AddItem ("ES")
    cboUF.AddItem ("GO")
    cboUF.AddItem ("MA")
    cboUF.AddItem ("MT")
    cboUF.AddItem ("MS")
    cboUF.AddItem ("MG")
    cboUF.AddItem ("PA")
    cboUF.AddItem ("PB")
    cboUF.AddItem ("PR")
    cboUF.AddItem ("PE")
    cboUF.AddItem ("PI")
    cboUF.AddItem ("RJ")
    cboUF.AddItem ("RN")
    cboUF.AddItem ("RS")
    cboUF.AddItem ("RO")
    cboUF.AddItem ("RR")
    cboUF.AddItem ("SC")
    cboUF.AddItem ("SP")
    cboUF.AddItem ("SE")
    cboUF.AddItem ("TO")
    
End Sub

Private Sub btnAtualizaCep_Click()
    frmAtualizaCep.Show 1
End Sub

Private Sub btnPesquisar_Click()
    Call DocXML(txtRua.Text, txtCidade.Text, cboUF.Text)

    lblInstrucaoDbClick.Visible = fxgDadosRuas.Rows > 1

End Sub

Public Function DocXML(rua As String, cidade As String, uf As String) As MSXML2.DOMDocument
    
    On Error GoTo trata_erro
    
    Dim objXML As MSXML2.ServerXMLHTTP
    Dim objXMLDoc As MSXML2.DOMDocument
    
    Set objXML = New MSXML2.ServerXMLHTTP
    Set objXMLDoc = New MSXML2.DOMDocument
    
    url = "https://viacep.com.br/ws/" & uf & "/" & cidade & "/" & rua & "/xml/"
    
    objXML.open "Get", url
    objXML.send
    
    If objXML.Status >= 400 And objXML.Status <= 505 Then
        MsgBox "Erro Ocorrido : " & objXML.Status & " - " & objXML.statusText, vbCritical, "Busca Cep"
        Exit Function
    Else
        objXMLDoc.loadXML (objXML.responseText)
    End If

    Dim objNodeList As IXMLDOMNodeList
    Set objNodeList = objXMLDoc.selectNodes("xmlcep/enderecos/endereco")
    
    
    If objNodeList.length > 0 Then
    
        With fxgDadosRuas
            .Clear
            .Cols = 6
            .Rows = objNodeList.length + 1
            
            Call MontaGrid(fxgDadosRuas)
                        
            For x = 0 To objNodeList.length - 1
            
                .TextMatrix(x + 1, 0) = objNodeList.Item(x).selectSingleNode("cep").Text
                .TextMatrix(x + 1, 1) = objNodeList.Item(x).selectSingleNode("logradouro").Text
                .TextMatrix(x + 1, 2) = objNodeList.Item(x).selectSingleNode("complemento").Text
                .TextMatrix(x + 1, 3) = objNodeList.Item(x).selectSingleNode("bairro").Text
                .TextMatrix(x + 1, 4) = objNodeList.Item(x).selectSingleNode("localidade").Text
                .TextMatrix(x + 1, 5) = objNodeList.Item(x).selectSingleNode("uf").Text
                
            Next
        End With

    Else
        MsgBox "Serviço indisponível - cep inválido"
    End If
    
    Set DocXML = objXMLDoc
    
    Exit Function
    
trata_erro:
    MsgBox Err.Description
    
End Function

Private Sub MontaGrid(grid As MSFlexGrid)
    With grid
        .TextMatrix(0, 0) = "Cep"
        .TextMatrix(0, 1) = "Rua"
        .TextMatrix(0, 2) = "Complemento"
        .TextMatrix(0, 3) = "Bairro"
        .TextMatrix(0, 4) = "Cidade"
        .TextMatrix(0, 5) = "UF"
        
        .ColWidth(0) = 1100
        .ColWidth(1) = 3600
        .ColWidth(2) = 2075
        .ColWidth(3) = 3600
        .ColWidth(4) = 3200
        .ColWidth(5) = 600
        
        For x = 0 To 5
            .Col = x
            .Row = 0
            .CellFontBold = True
        Next
        
        .ColAlignment(0) = AlignmentSettings.flexAlignCenterCenter
        .ColAlignment(5) = AlignmentSettings.flexAlignCenterCenter
        
    End With
End Sub

Private Sub fxgDadosRuas_DblClick()
    With fxgDadosRuas
        Clipboard.Clear
        Clipboard.SetText (.TextMatrix(.MouseRow, 0))
    End With
    frmAtualizaCep.Show 1
End Sub
