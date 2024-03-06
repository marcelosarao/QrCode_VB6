VERSION 5.00
Begin VB.Form FQrCode 
   Caption         =   "QrCode"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   17880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Limpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmd_GerarQrCode 
      Caption         =   "Gerar QrCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FQrCode.frx":0000
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o código que deseja gerar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   8895
      Left            =   6840
      Top             =   45
      Width           =   10995
   End
End
Attribute VB_Name = "FQrCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_GerarQrCode_Click()
    If Text1.Text <> "" Then
        Set Image1.Picture = QRCodegenBarcode(Text1.Text)
    End If
End Sub

Private Sub cmd_Limpar_Click()
    Text1 = ""
    Text1.SetFocus
End Sub

Private Sub cmd_GerarQrCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_GerarQrCode.ToolTipText = "Gera o código qr doque que foi digitado."
End Sub

Private Sub cmd_Limpar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_Limpar.ToolTipText = "Limpa o código qr doque que foi digitado, para fazer novo código."
End Sub

Private Sub Form_Load()
    Text1 = ""
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.ToolTipText = "Informe o código que deseja gerar."
End Sub
