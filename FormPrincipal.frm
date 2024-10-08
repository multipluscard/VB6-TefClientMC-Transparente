VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormularioPrincipal 
   BackColor       =   &H8000000E&
   Caption         =   "Exemplo TefClientMC - VB6"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   14385
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lbLog 
      Height          =   5520
      ItemData        =   "FormPrincipal.frx":0000
      Left            =   9240
      List            =   "FormPrincipal.frx":0002
      TabIndex        =   44
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton btnLimpaLog 
      BackColor       =   &H8000000E&
      Caption         =   "LIMPA LOG"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   11520
      TabIndex        =   37
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnAtributosAParte 
      BackColor       =   &H8000000E&
      Caption         =   "ATRIBUTOS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   12840
      TabIndex        =   36
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame panelTransacoes 
      BackColor       =   &H8000000E&
      Height          =   3495
      Left            =   9240
      TabIndex        =   31
      Top             =   6120
      Width           =   4935
      Begin VB.ListBox transacoespendentes 
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   4575
      End
      Begin VB.CommandButton btnDesfaz 
         Caption         =   "DESFAZ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   38
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CheckBox ckTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   2880
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CommandButton btnConfirma 
         Caption         =   "CONFIMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lbTransacoes 
         Caption         =   "TRANSAÇÕES"
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
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTabTipo 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483637
      TabCaption(0)   =   "CARTÃO"
      TabPicture(0)   =   "FormPrincipal.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LineCartao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tabTipoCartoes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "ADMINISTRAÇÃO"
      TabPicture(1)   =   "FormPrincipal.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LineAdm"
      Tab(1).Control(1)=   "btnExcluirBins"
      Tab(1).Control(2)=   "btnReimpressao"
      Tab(1).Control(3)=   "btnColetaCpf"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "QRMULTIPLUS"
      TabPicture(2)   =   "FormPrincipal.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LineQR"
      Tab(2).Control(1)=   "lbObs(1)"
      Tab(2).Control(2)=   "btnMenuPsp(0)"
      Tab(2).Control(3)=   "btnMenuPsp(1)"
      Tab(2).Control(4)=   "btnPspCliente"
      Tab(2).Control(5)=   "btnMercadoPago"
      Tab(2).Control(6)=   "btnPicPay"
      Tab(2).Control(7)=   "btnCancelarEstorno"
      Tab(2).Control(8)=   "btnStatusTransacao"
      Tab(2).ControlCount=   9
      Begin VB.CommandButton btnStatusTransacao 
         Caption         =   "STATUS TRANSAÇÃO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72360
         TabIndex        =   30
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton btnCancelarEstorno 
         Caption         =   "CANCELAR/ESTORNO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   29
         Top             =   4920
         Width           =   2175
      End
      Begin VB.CommandButton btnPicPay 
         Caption         =   "PICPAY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   28
         Top             =   3240
         Width           =   2925
      End
      Begin VB.CommandButton btnMercadoPago 
         Caption         =   "MERCADO PAGO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   27
         Top             =   2400
         Width           =   2925
      End
      Begin VB.CommandButton btnPspCliente 
         Caption         =   "PSP CLIENTE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   26
         Top             =   1560
         Width           =   2925
      End
      Begin VB.CommandButton btnMenuPsp 
         Caption         =   "MENU OPÇÕES PSP"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   -74640
         TabIndex        =   25
         Top             =   720
         Width           =   2925
      End
      Begin VB.CommandButton btnMenuPsp 
         Caption         =   "MENU OPÇÕES PSP"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   -74640
         TabIndex        =   24
         Top             =   720
         Width           =   2925
      End
      Begin VB.CommandButton btnColetaCpf 
         Caption         =   "COLETA DE CPF"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   23
         Top             =   2400
         Width           =   2925
      End
      Begin VB.CommandButton btnReimpressao 
         Caption         =   "REIMPRESSÃO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   22
         Top             =   1560
         Width           =   2925
      End
      Begin VB.CommandButton btnExcluirBins 
         Caption         =   "EXCLUIR BINS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   21
         Top             =   720
         Width           =   2925
      End
      Begin VB.CommandButton btnCancelar 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CANCELAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   5640
         Width           =   1815
      End
      Begin TabDlg.SSTab tabTipoCartoes 
         Height          =   4815
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "CREDITO"
         TabPicture(0)   =   "FormPrincipal.frx":0058
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "btnCancPreAutorizacao"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnConfPreAutorizacao(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "btnPreAutorizacao(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnConsulta(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "btnCreditoParceladoAdm"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "btnCreditoParceladoLoja(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "btnCreditoAVista(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "btnCredito(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "DEBITO"
         TabPicture(1)   =   "FormPrincipal.frx":0074
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "btnDebito"
         Tab(1).Control(1)=   "btnDebitoAVista"
         Tab(1).Control(2)=   "btnConsultaDebito"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "FROTA"
         TabPicture(2)   =   "FormPrincipal.frx":0090
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "btnFrota"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "VOUCHER"
         TabPicture(3)   =   "FormPrincipal.frx":00AC
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "btnVoucher"
         Tab(3).ControlCount=   1
         Begin VB.CommandButton btnConsultaDebito 
            Caption         =   "CONSULTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -74760
            TabIndex        =   20
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton btnDebitoAVista 
            Caption         =   "DÉBITO A VISTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74640
            TabIndex        =   19
            Top             =   1560
            Width           =   2925
         End
         Begin VB.CommandButton btnCredito 
            Caption         =   "CRÉDITO"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   360
            TabIndex        =   13
            Top             =   720
            Width           =   2925
         End
         Begin VB.CommandButton btnCreditoAVista 
            Caption         =   "CRÉDITO A VISTA"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   360
            TabIndex        =   12
            Top             =   1560
            Width           =   2925
         End
         Begin VB.CommandButton btnCreditoParceladoLoja 
            Caption         =   "CRÉDITO PARCELADO LOJA"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   360
            TabIndex        =   11
            Top             =   2400
            Width           =   2925
         End
         Begin VB.CommandButton btnCreditoParceladoAdm 
            Caption         =   "CRÉDITO PARCELADO ADM"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            TabIndex        =   10
            Top             =   3240
            Width           =   2925
         End
         Begin VB.CommandButton btnConsulta 
            Caption         =   "CONSULTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton btnPreAutorizacao 
            Caption         =   "PRÉ-AUTORIZAÇÃO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1560
            TabIndex        =   8
            Top             =   4200
            Width           =   1695
         End
         Begin VB.CommandButton btnConfPreAutorizacao 
            Caption         =   "CONF.PRÉ-AUTORIZAÇÃO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   3360
            TabIndex        =   7
            Top             =   4200
            Width           =   2295
         End
         Begin VB.CommandButton btnCancPreAutorizacao 
            Caption         =   "CANC.PRÉ-AUTORIZAÇÃO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5760
            TabIndex        =   6
            Top             =   4200
            Width           =   2295
         End
         Begin VB.CommandButton btnDebito 
            Caption         =   "DÉBITO"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74640
            TabIndex        =   5
            Top             =   720
            Width           =   2925
         End
         Begin VB.CommandButton btnFrota 
            Caption         =   "VENDA FROTA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74640
            TabIndex        =   4
            Top             =   720
            Width           =   2925
         End
         Begin VB.CommandButton btnVoucher 
            Caption         =   "VOUCHER"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74640
            TabIndex        =   3
            Top             =   720
            Width           =   2925
         End
      End
      Begin VB.Label lbObs 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Obs: Para visualizar o QR Code utilizar o parâmetro ExibirQrCode=2 no CliMC.ini"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   45
         Top             =   5640
         Width           =   3975
      End
      Begin VB.Line LineQR 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         X1              =   -74880
         X2              =   -66360
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line LineAdm 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   -74880
         X2              =   -66360
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line LineCartao 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         X1              =   240
         X2              =   8760
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame panelAtributos 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   9015
      Begin VB.TextBox txbValor 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txbNsu 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6720
         TabIndex        =   41
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txbCupom 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4560
         TabIndex        =   40
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txbParcela 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         TabIndex        =   39
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbNsu 
         BackColor       =   &H8000000E&
         Caption         =   "NSU"
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
         Left            =   6720
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbCupom 
         BackColor       =   &H8000000E&
         Caption         =   "CUPOM"
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
         Left            =   4560
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbParcela 
         BackColor       =   &H8000000E&
         Caption         =   "PARCELA"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label llValor 
         BackColor       =   &H8000000E&
         Caption         =   "VALOR"
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CheckBox ckMultiplus 
      BackColor       =   &H8000000E&
      Caption         =   " Multiplos Cartões"
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
      Left            =   7080
      TabIndex        =   35
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "FormularioPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FuncaoTeste Lib "ClientMC.dll" (ByVal Retorno As Byte) As Integer
Private Declare Function IniciaFuncaoMCInterativo Lib "TefClientMC.dll" (ByVal iComando As Integer, ByVal sCnpjCliente As String, ByVal iParcela As Integer, ByVal sCupom As String, ByVal sValor As String, ByVal sNsu As String, ByVal sData As String, ByVal sNumeroPDV As String, ByVal sCodigoLoja As String, ByVal sTipoComunicacao As Integer, ByVal sParametro As String) As Integer
Private Declare Function AguardaFuncaoMCInterativo Lib "TefClientMC.dll" () As String
Private Declare Function ContinuaFuncaoMCInterativo Lib "TefClientMC.dll" (ByVal sInformacao As String) As Long
Private Declare Function FinalizaFuncaoMCInterativo Lib "TefClientMC.dll" (ByVal iComando As Integer, ByVal sCnpjCliente As String, ByVal iParcela As Integer, ByVal sCupom As String, ByVal sValor As String, ByVal sNsu As String, ByVal sData As String, ByVal sNumeroPDV As String, ByVal sCodigoLoja As String, ByVal sTipoComunicacao As Integer, ByVal sParametro As String) As Integer
Private Declare Function CancelarFluxoMCInterativo Lib "TefClientMC.dll" () As Integer


Public Cnpj As String
Public pdv As String
Public codLoja As String
Public data As String
Public comunicacao As String
Public NovaVenda As Boolean

Dim parcela As String
Dim Cupom As String
Dim Nsu As String
Dim Valor As String
Dim Retorno As Integer

Dim confirmar As Boolean
Dim InsereNovoCartao As Boolean
Dim operacao As Integer

Dim lstCupons As New Collection
Dim Util As New Util

Function AdicionarLog(log As String)
   lbLog.AddItem log
   lbLog.ListIndex = lbLog.ListCount - 1
   lbLog.Refresh
End Function

Function ExecutarFuncao(operacao)
Dim retMsg As String
Dim arrMsg() As String
Dim strRetAguardaFMCInt As String
Dim retFim As Integer
Dim item As ListItem
Dim segString As String

confirmar = True

If (GetParametros(1)) Then
      'operacao = operacao
      Retorno = IniciaFuncaoMCInterativo(operacao, Cnpj, parcela, Cupom, Valor, Nsu, data, pdv, codLoja, comunicacao, "")

      Util.AdicionaLog "IniciaFuncaoMCInterativo()", ""
      
      If Retorno = 0 Then
            retMsg = ""
            While (retMsg <> "[ERROABORTAR]") And (retMsg <> "[RETORNO]") And (retMsg <> "[ERRODISPLAY]")
                  strRetAguardaFMCInt = AguardaFuncaoMCInterativo()
                  Util.AdicionaLog strRetAguardaFMCInt, ""
                  
                  If strRetAguardaFMCInt <> "" Then
                     AdicionarLog Format(Date, "dd/mm/yyyy") + " " + strRetAguardaFMCInt + vbNewLine
                     arrMsg = Split(strRetAguardaFMCInt, "#")
                     retMsg = arrMsg(0)
                     
                  Else
                     retMsg = ""
                  End If
  
                  If retMsg = "[MENU]" Then
                  
                     If (UBound(arrMsg) > 2) Then
                       answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                     Else
                       answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                     End If
                     
                     If LCase(answer) = "cancelar" Then
                        CancelarFluxoMCInterativo
                        Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                        MsgBox ("Fluxo Cancelado")
                        retMsg = "[ERROABORTAR]"
                        AdicionarLog Format(Date, "dd/mm/yyyy") + " - Fluxo Cancelado"
                        Util.AdicionaLog "Fluxo Cancelado", ""
                     Else
                        resp = ContinuaFuncaoMCInterativo(answer)
                     End If
                     
                  End If
                  
                  If retMsg = "[PERGUNTA]" Then
                    answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                     
                     If LCase(answer) = "cancelar" Then
                        CancelarFluxoMCInterativo
                        Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                        MsgBox ("Fluxo Cancelado")
                        retMsg = "[ERROABORTAR]"
                        AdicionarLog Format(Date, "dd/mm/yyyy") + " - Fluxo Cancelado"
                        Util.AdicionaLog "Fluxo Cancelado", ""
                     Else
                        resp = ContinuaFuncaoMCInterativo(answer)
                        
                     End If
                 
                  End If
                  
                 If retMsg = "[MSG]" Then
                     If (UBound(arrMsg) > 2) Then
                        If InStr(arrMsg(1), "SALDO") And InStr(arrMsg(1), "SALDO") Then
                           InputBox Replace(arrMsg(2), "|", vbNewLine), arrMsg(1)
                        End If
                     End If
                End If
                
                
            If retMsg = "[ERRODISPLAY]" Then
               answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
             
               If LCase(answer) = "cancelar" Then
                  CancelarFluxoMCInterativo
                  Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                  MsgBox ("Fluxo Cancelado")
                  retMsg = "[ERROABORTAR]"
                  AdicionarLog Format(Date, "dd/mm/yyyy") + " - Fluxo Cancelado"
                  Util.AdicionaLog "Fluxo Cancelado", ""
               Else
                  resp = ContinuaFuncaoMCInterativo(answer)
               End If
             
            End If
            
            Wend
            If retMsg = "[ERROABORTAR]" Then
               MsgBox (retMsg)
            End If
            If retMsg = "[RETORNO]" Then
            
               Dim strCupom As String
               Dim nsuRet As String
               Dim auxCupom As String
               
               strCupom = ""
               nsuRet = ""
               If (UBound(arrMsg) > 2) Then
                  If operacao > 50 And operacao < 60 Then
                     On Error GoTo ErrorHandler
                        Dim iFileNo As Integer
                        iFileNo = FreeFile
                        Open CStr(App.Path) + "\concluiupix.txt" For Output As #iFileNo
ErrorHandler:
                         MsgBox ex.Message
                         Util.AdicionaLog "Erro - " + ex.Message, ""
                  End If
                  
                  strCupom = Replace(arrMsg(15), "CAMPO122=", "")
                  strCupom = Replace(strCupom, "|", vbNewLine)
                  strCupom = Replace(strCupom, "CORTAR", vbNewLine + "-------------------------------" + vbNewLine)
                  
                  auxCupom = strCupom + "-------------------------------" + vbNewLine + vbNewLine
                  
                  nsuRet = Replace(arrMsg(5), "CAMPO0133=", "")
                  
               End If '>2
               If (UBound(arrMsg) = 2 And operacao = 200) Then
                  Dim aux() As String
                  auxCupom = Split(arrMsg(1), "=")
                  strCupom = "CPF Coletado: " + aux(1)
               End If '=2
               
               MsgBox (strCupom)
               MsgBox (Join(arrMsg, vbNewLine))
               
               retFim = 0
               If operacao <> 98 And operacao <> 99 Then
                  If confirmar Then
                     retFim = FinalizaFuncaoMCInterativo(98, Cnpj, parcela, Cupom, Valor, nsuRet, data, pdv, codLoja, comunicacao, "")
                  Else
                     retFim = FinalizaFuncaoMCInterativo(99, Cnpj, parcela, Cupom, Valor, nsuRet, data, pdv, codLoja, comunicacao, "")
                  End If
                   Util.AdicionaLog "FinalizaFuncaoMCInterativo()", ""
               Else
                  retFim = 0
               End If
               
               If retFim = 0 Then
                  AdicionarLog Format(Date, "dd/mm/yyyy") + " - FIM DA TRANSAÇÃO"
                  
               Else
                  AdicionarLog Format(Date, "dd/mm/yyyy") + " - ERRO: "
                  
               End If
               
            End If
      Else
         RetornaErro
         Util.AdicionaLog "Erro - IniciaFuncaoMCInterativo", ""
      End If
 Else
   MsgBox ("Parâmetros não conferem")
 End If
 
End Function

Function RetornaErro()
      If Retorno = 1 Then
          MsgBox "Erro genérico na execução", vbCritical, "Warning"
      ElseIf Retorno = 30 Then
          MsgBox "Não foi encontrado o caminho do ClientD.exe", vbCritical, "Warning"
      ElseIf Retorno = 31 Then
          MsgBox "ConfigMC.ini está vazio", vbCritical, "Warning"
      ElseIf Retorno = 32 Then
          MsgBox "ClientD.exe não encontrado", vbCritical, "Warning"
      ElseIf Retorno = 33 Then
          MsgBox "ClientD.exe não está em execução", vbCritical, "Warning"
      ElseIf Retorno = 34 Then
          MsgBox "Erro ao iniciar ClientD.exe", vbCritical, "Warning"
      ElseIf Retorno = 35 Then
          MsgBox "Não foi possível criar o arquivo de resposta", vbCritical, "Warning"
      ElseIf Retorno = 36 Then
          MsgBox "Erro na manipulação do arquivo de resposta", vbCritical, "Warning"
      ElseIf Retorno = 37 Then
          MsgBox "Erro na leitura do arquivo ConfigMC.ini", vbCritical, "Warning"
      ElseIf Retorno = 38 Then
          MsgBox "Valor da transação com formato incorreto", vbCritical, "Warning"
      ElseIf Retorno = 39 Then
          MsgBox "Executável de envio de transações não encontrado", vbCritical, "Warning"
      ElseIf Retorno = 40 Then
          MsgBox "CNPJ Inválido ou no formato incorreto", vbCritical, "Warning"
      ElseIf Retorno = 41 Then
          MsgBox "ClientD.exe está em processo de atualização", vbCritical, "Warning"
      ElseIf Retorno = 42 Then
          MsgBox "A automação não está sendo executada no modo administrador", vbCritical, "Warning"
      Else
          MsgBox Err.Description, vbCritical, "Warning"
      End If
End Function

Function GetParametros(valid As Integer)
   If Cnpj = "" Or pdv = "" Or codLoja = "" Or data = "" Or txbCupom.Text = "" Or txbNsu.Text = "" Or txbValor.Text = "" Or txbParcela.Text = "" Then
        answer = MsgBox("Informe todos os parâmetros", vbInformation, "Parâmetros")
        GetParametros = 0
   Else
        parcela = txbParcela.Text
        Nsu = txbNsu.Text
        Valor = txbValor.Text
        Cupom = txbCupom.Text
        GetParametros = 1
   End If
End Function

Private Sub AtribuirDados()
   Dim iFileNo As Integer
   iFileNo = FreeFile
   Dim sFileText As String
   Dim sFinal As String
   Dim arr() As String
   Dim aplicacao As String
     
   aplicacao = App.Path + "\ArqCarregaDados.txt"
   
   data = Format(Date, "yyyymmdd")
   comunicacao = 1
   
   If Dir(App.Path + "\ArqCarregaDados.txt", vbArchive) = "" Then Exit Sub
   
   Open aplicacao For Input As #iFileNo
   Do While Not EOF(iFileNo)
      sFileText = ""
      sFinal = ""
      Line Input #iFileNo, sFileText
      sFinal = sFinal & sFileText
      arr = Split(sFinal, "=")
     
      Select Case arr(0)
           Case "CNPJ"
                 Cnpj = CStr(arr(1))
           Case "PDV"
                 pdv = CStr(arr(1))
           Case "CODIGO_LOJA"
                 codLoja = CStr(arr(1))
      End Select
   Loop

End Sub

Private Sub btnAtributosAParte_Click(Index As Integer)
   FormularioAtributos.Show
End Sub

Private Sub btnCancelar_Click()
   ExecutarFuncao (5)
End Sub

Private Sub btnCancelarEstorno_Click()
ExecutarFuncao (54)
End Sub

Private Sub btnCancPreAutorizacao_Click()
   ExecutarFuncao (17)
End Sub

Private Sub btnColetaCpf_Click()
   ExecutarFuncao (200)
End Sub

Private Sub btnConfirma_Click(Index As Integer)
   Dim item As RetCupom
   Dim retFim As Integer
   Dim i As Integer
   Dim iTrans As Integer
   Dim transacao As RetCupom
   Dim sNsu As String
   
   If ckTodas.Value = 1 Then
      Dim auxParc As Integer
      Dim parc As Integer
      
         For i = 1 To lstCupons.Count()
         Set item = lstCupons.item(i)
         
         auxParc = CInt(item.Parcelas)
         If auxParc = 0 Then
            parc = 1
         Else
            parc = auxParc
         End If
         
         If i < lstCupons.Count() Then
         '.item(uoElem) Then 'ultimo elemento
            item.Nsu = item.Nsu + "*"
            Debug.Print item.Nsu
         End If
         
         retFim = FinalizaFuncaoMCInterativo(98, Cnpj, parc, item.Cupom, item.Valor, item.Nsu, data, pdv, codLoja, comunicacao, "")
         Util.AdicionaLog CStr(98) + Cnpj + CStr(item.Parcelas) + item.Cupom + item.Valor + sNsu + data + pdv + codLoja + CStr(comunicacao), ""
         Util.AdicionaLog "Retorno Confirmação/Desfazimento - " + CStr(98) + ":" + CStr(retFim), ""
         
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - Retorno Confirmação/Desfazimento - " + "98" + " : " + CStr(retFim)
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - " + "98" + ", " + CStr(Cnpj) + ", " + CStr(parc) + ", " + CStr(item.Cupom) + ", " + CStr(item.Valor) + ", " + CStr(item.Nsu) + ", " + CStr(data) + ", " + CStr(pdv) + ", " + codLoja + ", " + CStr(comunicacao)
         
      Next
      
      For i = 1 To lstCupons.Count()
         Set item = lstCupons.item(i)
         MsgBox item.Comprovante
      Next
      
      transacoespendentes.Clear
      AdicionarLog Format(Date, "dd/mm/yyyy") + " - FIM DA TRANSAÇÃO"
      
      ckMultiplus.Value = 0
   Else
   
      If transacoespendentes.ListIndex < 1 Then
         iTrans = 1
      Else
         iTrans = transacoespendentes.ListIndex + 1
      End If
      
      Set transacao = lstCupons.item(iTrans)
      
      MsgBox transacao.Comprovante
      
      auxParc = CInt(transacao.Parcelas)
      
      If auxParc = 0 Then
          parc = 1
      Else
          parc = auxParc
      End If
      
      sNsu = transacao.Nsu
      If lstCupons.Count > 1 Then
         sNsu = sNsu + "*"
      End If
      
      retFim = FinalizaFuncaoMCInterativo(98, Cnpj, CInt(parc), CStr(transacao.Cupom), transacao.Valor, sNsu, data, pdv, codLoja, comunicacao, "")
      
      Util.AdicionaLog CStr(98) + Cnpj + CStr(transacao.Parcelas) + transacao.Cupom + transacao.Valor + transacao.Nsu + data + pdv + codLoja + CStr(comunicacao), ""
      Util.AdicionaLog "Retorno Confirmação/Desfazimento - " + CStr(98) + ":" + CStr(retFim), ""
         
      AdicionarLog Format(Date, "dd/mm/yyyy") + " - Retorno Confirmação/Desfazimento - " + "98" + " : " + CStr(retFim)
      AdicionarLog Format(Date, "dd/mm/yyyy") + " - " + "98" + ", " + Cnpj + ", " + CStr(parc) + ", " + CStr(transacao.Cupom) + ", " + CStr(transacao.Valor) + ", " + CStr(transacao.Nsu) + ", " + data + ", " + pdv + ", " + codLoja + ", " + CStr(comunicacao)
   
      transacoespendentes.RemoveItem iTrans - 1
      transacoespendentes.Refresh
      
      If transacoespendentes.ListCount = 0 Then
         Set lstCupons = New Collection
      End If

      If lstCupons.Count = 0 Then
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - FIM DA TRANSAÇÃO"
         ckMultiplus.Value = 0
         Exit Sub
      End If
      
   End If
End Sub

Private Sub btnConfPreAutorizacao_Click(Index As Integer)
   ExecutarFuncao (16)
End Sub

Private Sub btnConsulta_Click(Index As Integer)
   ExecutarFuncao (9)
End Sub

Private Sub btnConsultaDebito_Click()
ExecutarFuncao (10)
End Sub

Private Sub btnCredito_Click(Index As Integer)
   operacao = 1
   ExecutarFuncao (operacao)
End Sub

Private Sub btnCreditoAVista_Click(Index As Integer)
   ExecutarFuncao (0)
End Sub

Private Sub btnCreditoParceladoAdm_Click()
   ExecutarFuncao (3)
End Sub

Private Sub btnCreditoParceladoLoja_Click(Index As Integer)
   ExecutarFuncao (2)
End Sub

Private Sub btnDebito_Click()
   ExecutarFuncao (4)
End Sub

Private Sub btnDebitoAVista_Click()
   ExecutarFuncao (20)
End Sub

Private Sub btnDesfaz_Click()

   If ckTodas.Value = 1 Then 'desfazer todas de uma vez
      Dim auxParc As Integer
      Dim parc As Integer
      
      For i = 1 To lstCupons.Count()
         Set item = lstCupons.item(i)
         
         auxParc = CInt(item.Parcelas)
         
         If auxParc = 0 Then
            parc = 1
         Else
            parc = auxParc
         End If
         
         If i < lstCupons.Count() Then
         '.item(uoElem) Then 'ultimo elemento
            item.Nsu = item.Nsu + "*"
            Debug.Print item.Nsu
         End If
         
         retFim = FinalizaFuncaoMCInterativo(99, Cnpj, parc, item.Cupom, item.Valor, item.Nsu, data, pdv, codLoja, comunicacao, "")
         
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - Retorno Confirmação/Desfazimento - " + "99" + " : " + CStr(retFim)
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - " + "99" + ", " + CStr(Cnpj) + ", " + CStr(parc) + ", " + CStr(item.Cupom) + ", " + CStr(item.Valor) + ", " + CStr(item.Nsu) + ", " + CStr(data) + ", " + CStr(pdv) + ", " + codLoja + ", " + CStr(comunicacao)
         
         Util.AdicionaLog CStr(99) + Cnpj + CStr(item.Parcelas) + item.Cupom + item.Valor + sNsu + data + pdv + codLoja + CStr(comunicacao), ""
         Util.AdicionaLog "Retorno Confirmação/Desfazimento - " + CStr(99) + ":" + CStr(retFim), ""
      
      Next
      
      For i = 1 To lstCupons.Count()
         Set item = lstCupons.item(i)
         MsgBox item.Comprovante
      Next
      
      transacoespendentes.Clear
      AdicionarLog Format(Date, "dd/mm/yyyy") + " - FIM DA TRANSAÇÃO"
      
      ckMultiplus.Value = 0
         
   Else 'desfazer uma de cada vez
      If transacoespendentes.ListIndex < 1 Then
         iTrans = 1
      Else
         iTrans = transacoespendentes.ListIndex + 1
      End If
      
      Set transacao = lstCupons.item(iTrans)
      
      MsgBox transacao.Comprovante
      
      auxParc = CInt(transacao.Parcelas)
      
      If auxParc = 0 Then
          parc = 1
      Else
          parc = auxParc
      End If
      
      sNsu = transacao.Nsu
      If lstCupons.Count > 1 Then
         sNsu = sNsu + "*"
      End If
      
      retFim = FinalizaFuncaoMCInterativo(99, Cnpj, CInt(parc), CStr(transacao.Cupom), transacao.Valor, sNsu, data, pdv, codLoja, comunicacao, "")
      
      Util.AdicionaLog CStr(98) + Cnpj + CStr(transacao.Parcelas) + transacao.Cupom + transacao.Valor + transacao.Nsu + data + pdv + codLoja + CStr(comunicacao), ""
      Util.AdicionaLog "Retorno Confirmação/Desfazimento - " + CStr(98) + ":" + CStr(retFim), ""
      
      AdicionarLog Format(Date, "dd/mm/yyyy") + " - Retorno Confirmação/Desfazimento - " + "99" + " : " + CStr(retFim)
      AdicionarLog Format(Date, "dd/mm/yyyy") + " - " + "99" + ", " + Cnpj + ", " + CStr(parc) + ", " + CStr(transacao.Cupom) + ", " + CStr(transacao.Valor) + ", " + CStr(transacao.Nsu) + ", " + data + ", " + pdv + ", " + codLoja + ", " + CStr(comunicacao)
   
      transacoespendentes.RemoveItem iTrans - 1
      transacoespendentes.Refresh
      
      If transacoespendentes.ListCount = 0 Then
         Set lstCupons = New Collection
      End If
      
      If lstCupons.Count = 0 Then
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - FIM DA TRANSAÇÃO"
         ckMultiplus.Value = 0
      End If
      
   End If
   
End Sub

Private Sub btnExcluirBins_Click()
   Dim ret As Integer
   If GetParametros(1) Then
      ret = IniciaFuncaoMCInterativo(110, Cnpj, parcela, Cupom, Valor, Nsu, data, pdv, codLoja, comunicacao, "")
   Else
      MsgBox Err.Description
   End If
End Sub

Private Sub btnFrota_Click()
ExecutarFuncao (11)
End Sub

Private Sub btnLimpaLog_Click(Index As Integer)
   lbLog.Clear
End Sub



Private Sub btnMenuPsp_Click(Index As Integer)
   ExecutarFuncao (50)
End Sub

Private Sub btnMercadoPago_Click()
ExecutarFuncao (52)
End Sub

Private Sub btnPicPay_Click()
ExecutarFuncao (53)
End Sub

Private Sub btnPreAutorizacao_Click(Index As Integer)
ExecutarFuncao (15)
End Sub

Private Sub btnPspCliente_Click()
ExecutarFuncao (51)
End Sub

Private Sub btnReimpressao_Click()
   ExecutarFuncao (6)
End Sub

Private Sub btnStatusTransacao_Click()
ExecutarFuncao (56)
End Sub

Private Sub btnVoucher_Click()
ExecutarFuncao (18)
End Sub

Private Sub ckMultiplus_Click()
Dim retMsg As String
Dim arrMsg() As String
Dim strRetAguardaFMCInt As String
Dim vRetCupom As String
Dim dadosCupom As RetCupom
   NovaVenda = True
   
   If GetParametros(1) And ckMultiplus.Value = 1 Then
      While NovaVenda = True
         panelTransacoes.Visible = True
         lbLog.Height = 5325
      
         If InsereNovoCartao <> True Then
            Set lstCupons = New Collection
            'Set lstExibicao = New Collection
         End If
         
         Nsu = Nsu + "*"
         Retorno = IniciaFuncaoMCInterativo(operacao, Cnpj, parcela, Cupom, Valor, Nsu, data, pdv, codLoja, comunicacao, "")
         AdicionarLog Format(Date, "dd/mm/yyyy") + " - IniciaFuncaoMCInterativo()" + vbNewLine
         Util.AdicionaLog "IniciaFuncaoMCInterativo()", ""
         
         If Retorno = 0 Then
               retMsg = ""
               While (retMsg <> "[ERROABORTAR]") And (retMsg <> "[RETORNO]") And (retMsg <> "[ERRODISPLAY]")
                  strRetAguardaFMCInt = AguardaFuncaoMCInterativo()
                  Util.AdicionaLog strRetAguardaFMCInt, ""
                  
                  If strRetAguardaFMCInt <> "" Then
                     AdicionarLog Format(Date, "dd/mm/yyyy") + " " + strRetAguardaFMCInt + vbNewLine
                     arrMsg = Split(strRetAguardaFMCInt, "#")
                     retMsg = arrMsg(0)
                     
                  Else
                     retMsg = ""
                  End If
                  
                  If retMsg = "[MENU]" Then
                     If (UBound(arrMsg) > 2) Then
                       answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                     Else
                       answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                     End If
                     
                     If LCase(answer) = "cancelar" Then
                        CancelarFluxoMCInterativo
                        Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                        MsgBox ("Fluxo Cancelado")
                        retMsg = "[ERROABORTAR]"
                        AdicionarLog Format(Date, "dd/mm/yyyy") + " - Fluxo Cancelado"
                        Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                        NovaVenda = False
                     Else
                        resp = ContinuaFuncaoMCInterativo(answer)
                        
                     End If
                  End If '[MENU]
                  
                 If retMsg = "[PERGUNTA]" Then
                   answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                   
                   If LCase(answer) = "cancelar" Then
                      CancelarFluxoMCInterativo
                      Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                      MsgBox ("Fluxo Cancelado")
                      retMsg = "[ERROABORTAR]"
                      AdicionarLog Format(Date, "dd/mm/yyyy") + " - Fluxo Cancelado"
                      Util.AdicionaLog "Fluxo Cancelado", ""
                      NovaVenda = False
                   Else
                      resp = ContinuaFuncaoMCInterativo(answer)
                   End If
                End If '[PERGUNTA]
                
               If retMsg = "[MSG]" Then
                    If (UBound(arrMsg) > 2) Then
                       If InStr(arrMsg(1), "SALDO") And InStr(arrMsg(1), "SALDO") Then
                          InputBox Replace(arrMsg(2), "|", vbNewLine), arrMsg(1)
                       End If
                    End If
               End If '[MSG]
               
               If retMsg = "[ERRODISPLAY]" Then
                  answer = InputBox(Replace(arrMsg(2), "|", vbNewLine), arrMsg(1))
                
                  If LCase(answer) = "cancelar" Then
                        CancelarFluxoMCInterativo
                        Util.AdicionaLog "CancelarFluxoMCInterativo()", ""
                        MsgBox ("Fluxo Cancelado")
                        retMsg = "[ERROABORTAR]"
                        AdicionarLog Format(Date, "dd/mm/yyyy") + " - Fluxo Cancelado"
                        Util.AdicionaLog "Fluxo Cancelado", ""
                        NovaVenda = False
                     Else
                        resp = ContinuaFuncaoMCInterativo(answer)
                     End If
               End If '[ERRODISPLAY]
               Wend 'retMsg <> "[ERROABORTAR]"
               
               If retMsg = "[ERROABORTAR]" Then
                  MsgBox (retMsg)
               End If
               If retMsg = "[RETORNO]" Then
                  Dim strCupom As String
                  Dim auxCupom As String
                  
                  strCupom = Replace(arrMsg(15), "CAMPO122=", "")
                  strCupom = Replace(strCupom, "|", vbNewLine)
                  strCupom = Replace(strCupom, "CORTAR", vbNewLine + "-------------------------------" + vbNewLine)
                  auxCupom = strCupom + "-------------------------------" + vbNewLine + vbNewLine
                  
                  Set dadosCupom = New RetCupom
                  
                  dadosCupom.Comprovante = Replace(arrMsg(15), "CAMPO122=", "")
                  dadosCupom.Comprovante = Replace(dadosCupom.Comprovante, "|", vbNewLine)
                  dadosCupom.Comprovante = Replace(dadosCupom.Comprovante, "CORTAR", vbNewLine + "-------------------------------" + vbNewLine)
                  dadosCupom.VenctoCartao = Replace(arrMsg(14), "CAMPO0513=", "")
                  dadosCupom.NsuRede = Replace(arrMsg(13), "CAMPO0134=", "")
                  dadosCupom.Cliente = Replace(arrMsg(12), "CAMPO1003=", "")
                  dadosCupom.Cnpj = Replace(arrMsg(11), "CAMPO0950=", "")
                  dadosCupom.UltimosDigitos = Replace(arrMsg(10), "CAMPO1190=", "")
                  dadosCupom.BinCartao = Replace(arrMsg(9), "CAMPO0136=", "")
                  dadosCupom.TxServico = Replace(arrMsg(8), "CAMPO0504=", "")
                  dadosCupom.Parcelas = Replace(arrMsg(7), "CAMPO0505=", "")
                  dadosCupom.Nsu = Replace(arrMsg(6), "CAMPO0133=", "")
                  dadosCupom.CodAutorizacao = Replace(arrMsg(5), "CAMPO0135=", "")
                  dadosCupom.CodRede = Replace(arrMsg(4), "CAMPO0131=", "")
                  dadosCupom.CodBandeira = Replace(arrMsg(3), "CAMPO0132=", "")
                  dadosCupom.Valor = Replace(arrMsg(2), "CAMPO0002=", "")
                  dadosCupom.Cupom = Replace(arrMsg(1), "CAMPO0160=", "")
                
                  lstCupons.Add dadosCupom
                  transacoespendentes.AddItem Replace(dadosCupom.Comprovante, vbNewLine, "|"), 0
                  transacoespendentes.Refresh
               End If
            
         Else 'Else Retorno = 0
            MsgBox "Erro - IniciaFuncaoMCInterativo"
            Util.AdicionaLog "Erro - IniciaFuncaoMCInterativo", ""
         End If 'If retorno = 0

         If NovaVenda = True Then
            If MsgBox("Deseja Pagar com mais um cartão", vbYesNo, "Pagar com Multiplus Cartões") = vbYes Then
               Dim novoValor As String
               InsereNovoCartao = True
               novoValor = InputBox("Digite o valor da Transação:", "Multiplus Cartões")
               Valor = novoValor
            Else
               InsereNovoCartao = False
               NovaVenda = False
               Exit Sub
            End If
         Else
            ckMultiplus.Value = 0
            panelTransacoes.Visible = False
         End If 'NovaVenda = True
         
      Wend 'WHILE NOVAVENDA
   Else 'IF GetParametros(1) And ckMultiplus.Value = 1
      NovaVenda = False
      ckMultiplus.Value = 0
   End If
      

End Sub

Private Sub Form_Load()
   operacao = 0
   NovaVenda = True
   If ckMultiplus.Value = 1 Then
      panelTransacoes.Visible = True
      lbLog.Height = 5325
    Else
      panelTransacoes.Visible = False
      lbLog.Height = 9030
    End If
   
    AtribuirDados
    'ChDir "C:\DLL" 'Se necessário Mudar o diretório atual para reconhecer a dll
   
End Sub


