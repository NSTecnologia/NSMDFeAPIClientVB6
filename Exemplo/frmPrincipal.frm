VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbTpDown 
      Height          =   315
      ItemData        =   "frmPrincipal.frx":0000
      Left            =   6480
      List            =   "frmPrincipal.frx":0013
      TabIndex        =   29
      Text            =   "XP"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   8760
      TabIndex        =   28
      Top             =   960
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtTpAmb 
      Height          =   315
      Left            =   6480
      TabIndex        =   13
      Text            =   "2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtCaminho 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "C:/Notas/"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "07364617000135"
      Top             =   1440
      Width           =   3495
   End
   Begin TabDlg.SSTab tabControl 
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Emissão"
      TabPicture(0)   =   "frmPrincipal.frx":0028
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cbTpConteudo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtRetornoEmissao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtConteudo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEnviar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Encerramento"
      TabPicture(1)   =   "frmPrincipal.frx":0044
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtProtEnc"
      Tab(1).Control(1)=   "cmdEncerramento"
      Tab(1).Control(2)=   "txtRetornoEnc"
      Tab(1).Control(3)=   "txtCodMun"
      Tab(1).Control(4)=   "txtCodUF"
      Tab(1).Control(5)=   "txtDHEvento"
      Tab(1).Control(6)=   "txtDTEnc"
      Tab(1).Control(7)=   "txtChaveEnc"
      Tab(1).Control(8)=   "Label12"
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(11)=   "Label9"
      Tab(1).Control(12)=   "Label8"
      Tab(1).Control(13)=   "Label7"
      Tab(1).Control(14)=   "Label6"
      Tab(1).ControlCount=   15
      Begin VB.TextBox txtProtEnc 
         Height          =   315
         Left            =   -69600
         TabIndex        =   26
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton cmdEncerramento 
         Caption         =   "Encerramento >>>>>>"
         Height          =   615
         Left            =   -68640
         TabIndex        =   25
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtRetornoEnc 
         Height          =   3015
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   4200
         Width           =   10215
      End
      Begin VB.TextBox txtCodMun 
         Height          =   315
         Left            =   -71640
         TabIndex        =   21
         Text            =   "4303509"
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtCodUF 
         Height          =   315
         Left            =   -71640
         TabIndex        =   19
         Text            =   "43"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtDHEvento 
         Height          =   315
         Left            =   -74880
         TabIndex        =   17
         Text            =   "2019-04-12T11:40:00-03:00"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtDTEnc 
         Height          =   315
         Left            =   -74880
         TabIndex        =   15
         Text            =   "2019-04-12"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtChaveEnc 
         Height          =   315
         Left            =   -74880
         TabIndex        =   11
         Top             =   840
         Width           =   5175
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Enviar Documento para Processamento >>>>>>"
         Height          =   615
         Left            =   6600
         TabIndex        =   8
         Top             =   3840
         Width           =   3735
      End
      Begin VB.TextBox txtConteudo 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   840
         Width           =   10215
      End
      Begin VB.TextBox txtRetornoEmissao 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   4800
         Width           =   10215
      End
      Begin VB.ComboBox cbTpConteudo 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":0060
         Left            =   8280
         List            =   "frmPrincipal.frx":006D
         TabIndex        =   5
         Text            =   "txt"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   27
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Resposta da API:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   24
         Top             =   3840
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Código do município:"
         Height          =   195
         Left            =   -71640
         TabIndex        =   22
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Código UF:"
         Height          =   195
         Left            =   -71640
         TabIndex        =   20
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Data-hora evento:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   18
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data encerramento"
         Height          =   195
         Left            =   -74880
         TabIndex        =   16
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Chave:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conteudo"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Resposta da API:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   4560
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Ambiente:"
      Height          =   195
      Left            =   6480
      TabIndex        =   14
      Top             =   360
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salvar em:"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   750
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()
    On Error GoTo SAI
    Dim resposta As String
    Dim statusEnvio, statusConsulta, statusDownload, cStat, chMDFe, nProt, motivo, nsNRec, erros As String
    
    If (txtCaminho.Text <> "") And (txtConteudo.Text <> "") And (cbTpConteudo.Text <> "") And (cbTpDown.Text <> "") And (txtTpAmb.Text <> "") Then
        resposta = emitirMDFeSincrono(txtConteudo.Text, cbTpConteudo.Text, txtCNPJ.Text, cbTpDown.Text, txtTpAmb.Text, txtCaminho.Text, checkExibir.Value)
        txtChaveEnc.Text = LerDadosJSON(resposta, "chMDFe", "", "")
        txtProtEnc.Text = LerDadosJSON(resposta, "nProt", "", "")
        txtRetornoEmissao.Text = resposta
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
 
    'Tratamento de retorno
    statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
    statusConsulta = LerDadosJSON(retorno, "statusConsulta", "", "")
    statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
    cStat = LerDadosJSON(retorno, "cStat", "", "")
    chMDFe = LerDadosJSON(retorno, "chMDFe", "", "")
    nProt = LerDadosJSON(retorno, "nProt", "", "")
    motivo = LerDadosJSON(retorno, "motivo", "", "")
    nsNRec = LerDadosJSON(retorno, "nsNRec", "", "")
    erros = LerDadosJSON(retorno, "erros", "", "")
    
    'Agora que você já leu os dados, é aconselhável que faça o salvamento de todos
    'eles no seu banco de dados antes de prosseguir para o teste abaixo
             
    If (statusEnvio = 200) Or (statusEnvio = -6) Then
  
        If (statusConsulta = 200) Then
            
            If (cStat = 100) Then
                MsgBox (motivo)
                 
                If (statusDownload <> 200) Then
                    MsgBox ("Erro ao fazer o Download")
                End If
            Else
                MsgBox (motivo)
            End If
        Else
            MsgBox (motivo + Chr(13) + erros)
        End If
    Else
        MsgBox (motivo + Chr(13) + erros)
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI

End Sub

Private Sub cmdEncerramento_Click()
    On Error GoTo SAI
    Dim resposta As String
    
    If (txtChaveEnc.Text <> "") And (txtProtEnc.Text <> "") And (txtDHEvento.Text <> "") And (txtDTEnc.Text <> "") And (txtCodUF.Text <> "") And (txtCodMun.Text <> "") And (cbTpDown.Text <> "") And (txtTpAmb.Text <> "") Then
        resposta = encerrarMDFe(txtChaveEnc.Text, txtProtEnc.Text, txtDHEvento.Text, txtDTEnc.Text, txtCodUF.Text, txtCodMun.Text, txtTpAmb.Text, cbTpDown.Text, txtCaminho.Text, checkExibir.Value)
        txtRetornoEnc.Text = resposta
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI

End Sub
