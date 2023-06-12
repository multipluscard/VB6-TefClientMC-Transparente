VERSION 5.00
Begin VB.Form FormularioAtributos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Formulario Atributos"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleMode       =   0  'User
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame panelAtributos 
      BackColor       =   &H8000000E&
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txbPdv 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txbData 
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox txbCodLoja 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txbCnpj 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txbComunicacao 
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lbPdv 
         BackColor       =   &H8000000E&
         Caption         =   "PDV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lbData 
         BackColor       =   &H8000000E&
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lbCodLoja 
         BackColor       =   &H8000000E&
         Caption         =   "COD LOJA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lbCnpj 
         BackColor       =   &H8000000E&
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lbComunicaco 
         BackColor       =   &H8000000E&
         Caption         =   "COMUNICAÇÃO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FormularioAtributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   txbData.Text = FormularioPrincipal.data
   txbCnpj.Text = FormularioPrincipal.cnpj
   txbCodLoja.Text = FormularioPrincipal.codLoja
   txbComunicacao.Text = FormularioPrincipal.comunicacao
   txbPdv.Text = FormularioPrincipal.pdv
End Sub

Private Sub txbComunicacao_GotFocus()
   On Error Resume Next
   txbComunicacao.SelStart = 0
   txbComunicacao.SelLength = Len(txbComunicacao.Text)
End Sub
Private Sub txbComunicacao_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 FormularioPrincipal.comunicacao = txbComunicacao.Text
End Sub

Private Sub txbData_GotFocus()
   On Error Resume Next
   txbData.SelStart = 0
   txbData.SelLength = Len(txbData.Text)
End Sub

Private Sub txbData_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 FormularioPrincipal.data = txbData.Text
End Sub

Private Sub txbCnpj_GotFocus()
   On Error Resume Next
   txbCnpj.SelStart = 0
   txbCnpj.SelLength = Len(txbCnpj.Text)
End Sub

Private Sub txbCnpj_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 FormularioPrincipal.cnpj = txbCnpj.Text
End Sub

Private Sub txbCodLoja_GotFocus()
   On Error Resume Next
   txbCodLoja.SelStart = 0
   txbCodLoja.SelLength = Len(txbCodLoja.Text)
End Sub

Private Sub txbCodLoja_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 FormularioPrincipal.codLoja = txbCodLoja.Text
End Sub

Private Sub txbPdv_GotFocus()
   On Error Resume Next
   txbPdv.SelStart = 0
   txbPdv.SelLength = Len(txbPdv.Text)
End Sub

Private Sub txbPdv_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 FormularioPrincipal.pdv = txbPdv.Text
End Sub

