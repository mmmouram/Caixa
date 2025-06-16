
VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetalhe_Pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido - Detalhe"
   ClientHeight    =   7095
   ClientLeft      =   1905
   ClientTop       =   1860
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCLIENTE 
      Caption         =   "Cliente"
      Height          =   680
      Left            =   2280
      TabIndex        =   14
      Top             =   240
      Width           =   8055
      Begin VB.TextBox txtCNPJ 
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1770
      End
      Begin VB.TextBox txtRAZAO_SOCIAL 
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   5010
      End
      Begin VB.Label lblLinha 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   2460
         TabIndex        =   18
         Top             =   285
         Width           =   105
      End
      Begin VB.Label lblCNPJ 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ : "
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   285
         Width           =   525
      End
   End
   Begin MSComctlLib.ImageList ImgPretoBranco 
      Left            =   9840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":08DA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":11B4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":1A8E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":2368
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":2C42
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":351E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":3DFA
            Key             =   "IMG8"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgColorido 
      Left            =   10560
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":46D6
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":4FB0
            Key             =   "Propriedade"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":588A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":6164
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":6A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":7318
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":7BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalhe_Pedido.frx":84D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11700
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   600
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   330
         Width           =   1095
      End
      Begin VB.Frame fraPEDIDO 
         Caption         =   "Pedido"
         Height          =   680
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtNUM_PEDIDO 
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblCODIGO 
            AutoSize        =   -1  'True
            Caption         =   "N mero : "
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   285
            Width           =   705
         End
      End
      Begin VB.Frame fraFUNDO1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6120
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   11490
         Begin VB.Frame fraNOTA_FISCAL 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5175
            Left            =   1380
            TabIndex        =   19
            Top             =   600
            Width           =   10995
            Begin VB.CommandButton cmdAtualizar_Nota_Fiscal 
               Caption         =   "&Atualizar"
               Height          =   315
               Left            =   9442
               TabIndex        =   40
               Top             =   250
               Width           =   870
            End
            Begin VB.CommandButton cmdDetalhar_Nota_Fiscal 
               Caption         =   "&Detalhar"
               Height          =   315
               Left            =   8490
               TabIndex        =   39
               Top             =   250
               Width           =   870
            End
            Begin VB.CommandButton cmdExcel_Nota_Fiscal 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10395
               Picture         =   "frmDetalhe_Pedido.frx":8DAC
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   220
               Width           =   495
            End
            Begin VB.CheckBox chkCARREGAR_NOTA_FISCAL_PEDIDO 
               Caption         =   "Carregar Automaticamente"
               Height          =   240
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Value           =   1  'Checked
               Width           =   2775
            End
            Begin VB.Frame fra_lvwNota_Fical 
               Caption         =   "Listagem Geral"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4395
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   10770
               Begin MSComctlLib.ListView lvwNota_Fiscal_Pedido 
                  Height          =   4005
                  Left            =   135
                  TabIndex        =   21
                  Top             =   240
                  Width           =   10500
                  _ExtentX        =   18521
                  _ExtentY        =   7064
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  HotTracking     =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   0
               End
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nota Fiscal"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   360
               Index           =   6
               Left            =   4665
               TabIndex        =   24
               Top             =   210
               Width           =   1530
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nota Fiscal"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Index           =   7
               Left            =   4680
               TabIndex        =   25
               Top             =   210
               Width           =   1530
            End
         End
         Begin VB.Frame fraPEDIDO_BLOQUEIOS 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5175
            Left            =   960
            TabIndex        =   26
            Top             =   600
            Width           =   10995
            Begin VB.CommandButton cmdAtualizar_Pedido_Bloqueios 
               Caption         =   "&Atualizar"
               Height          =   315
               Left            =   9442
               TabIndex        =   43
               Top             =   250
               Width           =   870
            End
            Begin VB.Frame fra_lvw_Pedido_Bloqueios 
               Caption         =   "Listagem Geral"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4395
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   10770
               Begin MSComctlLib.ListView lvwPedido_Bloqueios_Pedido 
                  Height          =   4005
                  Left            =   135
                  TabIndex        =   30
                  Top             =   240
                  Width           =   10500
                  _ExtentX        =   18521
                  _ExtentY        =   7064
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  HotTracking     =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   0
               End
            End
            Begin VB.CheckBox chkCARREGAR_PEDIDO_BLOQUEIOS_PEDIDO 
               Caption         =   "Carregar Automaticamente"
               Height          =   240
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Value           =   1  'Checked
               Width           =   2775
            End
            Begin VB.CommandButton cmdExcel_Pedido_Bloqueios 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10395
               Picture         =   "frmDetalhe_Pedido.frx":9132
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   220
               Width           =   495
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido Bloqueios"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   360
               Index           =   4
               Left            =   4665
               TabIndex        =   31
               Top             =   210
               Width           =   2310
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido Bloqueios"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Index           =   5
               Left            =   4680
               TabIndex        =   32
               Top             =   210
               Width           =   2310
            End
         End
         Begin VB.Frame fraOBSERVACAO_PEDIDO 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5175
            Left            =   570
            TabIndex        =   33
            Top             =   600
            Width           =   10995
            Begin VB.CommandButton cmdAtualizar_Observacao_Pedido 
               Caption         =   "&Atualizar"
               Height          =   315
               Left            =   9442
               TabIndex        =   42
               Top             =   250
               Width           =   870
            End
            Begin VB.CommandButton cmdExcel_Observacao_Pedido 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10395
               Picture         =   "frmDetalhe_Pedido.frx":94B8
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   220
               Width           =   495
            End
            Begin VB.CheckBox chkCARREGAR_OBSERVACAO_PEDIDO_PEDIDO 
               Caption         =   "Carregar Automaticamente"
               Height          =   240
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Value           =   1  'Checked
               Width           =   2775
            End
            Begin VB.Frame fra_lvwObservacao_Pedido 
               Caption         =   "Listagem Geral"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4395
               Left            =   120
               TabIndex        =   34
               Top             =   600
               Width           =   10770
               Begin MSComctlLib.ListView lvwObservacao_Pedido_Pedido 
                  Height          =   4005
                  Left            =   135
                  TabIndex        =   35
                  Top             =   240
                  Width           =   10500
                  _ExtentX        =   18521
                  _ExtentY        =   7064
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  HotTracking     =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   0
               End
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observa  o Pedido"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   360
               Index           =   2
               Left            =   4665
               TabIndex        =   37
               Top             =   210
               Width           =   2610
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observa  o Pedido"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Index           =   3
               Left            =   4680
               TabIndex        =   38
               Top             =   210
               Width           =   2610
            End
         End
         Begin VB.Frame fraITENS_PEDIDO 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5175
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   10995
            Begin VB.CommandButton cmdAtualizar_Itens_Pedido 
               Caption         =   "&Atualizar"
               Height          =   315
               Left            =   9442
               TabIndex        =   44
               Top             =   250
               Width           =   870
            End
            Begin VB.CommandButton cmdExcel_Itens_Pedido 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10395
               Picture         =   "frmDetalhe_Pedido.frx":983E
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   220
               Width           =   495
            End
            Begin VB.CheckBox chkCARREGAR_ITENS_PEDIDO_PEDIDO 
               Caption         =   "Carregar Automaticamente"
               Height          =   240
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Value           =   1  'Checked
               Width           =   2775
            End
            Begin VB.Frame fra_lvwItens_Pedido 
               Caption         =   "Listagem Geral"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4395
               Left            =   120
               TabIndex        =   5
               Top             =   600
               Width           =   10770
               Begin MSComctlLib.ListView lvwItens_Pedido_Pedido 
                  Height          =   4005
                  Left            =   135
                  TabIndex        =   6
                  Top             =   240
                  Width           =   10500
                  _ExtentX        =   18521
                  _ExtentY        =   7064
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  HotTracking     =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   0
               End
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Itens Pedido"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   360
               Index           =   0
               Left            =   4665
               TabIndex        =   7
               Top             =   210
               Width           =   1650
            End
            Begin VB.Label lblTituloForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Itens Pedido"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Index           =   1
               Left            =   4680
               TabIndex        =   8
               Top             =   210
               Width           =   1650
            End
         End
         Begin MSComctlLib.TabStrip tabDetalhe 
            Height          =   5775
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   11220
            _ExtentX        =   19791
            _ExtentY        =   10186
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   4
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Itens"
                  Key             =   "Itens_Pedido"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Observa  o"
                  Key             =   "Observacao_Pedido"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Pedido Bloqueios"
                  Key             =   "Pedido_Bloqueios"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Nota Fiscal"
                  Key             =   "Nota_Fiscal"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmDetalhe_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gstrCod              As String
Private gstrCNPJ             As String
Private gstrRAZAO_SOCIAL     As String
Public Property Get Codigo() As String
   Codigo = gstrCod
End Property

Public Property Let Codigo(pCOD As String)
   gstrCod = pCOD
End Property

Public Property Get CNPJ() As String
   CNPJ = gstrCNPJ
End Property

Public Property Let CNPJ(pCNPJ As String)
   gstrCNPJ = pCNPJ
End Property
Public Property Get RAZAO_SOCIAL() As String
   RAZAO_SOCIAL = gstrRAZAO_SOCIAL
End Property

Public Property Let RAZAO_SOCIAL(pRAZAO_SOCIAL As String)
   gstrRAZAO_SOCIAL = pRAZAO_SOCIAL
End Property

Private Sub cmdAtualizar_Itens_Pedido_Click()
    Atualiza_Lista_Itens_Pedido
End Sub

Private Sub cmdAtualizar_Nota_Fiscal_Click()
    Atualiza_Lista_Nota_Fiscal
End Sub

Private Sub cmdAtualizar_Observacao_Pedido_Click()
    Atualiza_Lista_Observacao_Pedido
End Sub

Private Sub cmdAtualizar_Pedido_Bloqueios_Click()
    Atualiza_Lista_Pedido_Bloqueios
End Sub

Private Sub cmdDetalhar_Nota_Fiscal_Click()
     lvwNota_Fiscal_Pedido_DblClick
End Sub

Private Sub cmdExcel_Itens_Pedido_Click()
    GerarExcel_ListView lvwItens_Pedido_Pedido
End Sub

Private Sub cmdExcel_Observacao_Pedido_Click()
    GerarExcel_ListView lvwObservacao_Pedido_Pedido
End Sub

Private Sub cmdExcel_Pedido_Bloqueios_Click()
    GerarExcel_ListView lvwPedido_Bloqueios_Pedido
End Sub
Private Sub cmdExcel_Nota_Fiscal_Click()
    GerarExcel_ListView lvwNota_Fiscal_Pedido
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
   Screen.MousePointer = vbHourglass
   
   PreparaForm Me
      
   Centra_Form Me, False
   
   RetornaCheckbox chkCARREGAR_ITENS_PEDIDO_PEDIDO
   RetornaCheckbox chkCARREGAR_OBSERVACAO_PEDIDO_PEDIDO
   RetornaCheckbox chkCARREGAR_PEDIDO_BLOQUEIOS_PEDIDO
   RetornaCheckbox chkCARREGAR_NOTA_FISCAL_PEDIDO
   
   Atualiza_Controles
   
   tabDetalhe_Click
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Atualiza_Controles()
   
   '---- Itens
   If chkCARREGAR_ITENS_PEDIDO_PEDIDO.Value = vbChecked Then
        Atualiza_Lista_Itens_Pedido
   End If
   
   If chkCARREGAR_OBSERVACAO_PEDIDO_PEDIDO.Value = vbChecked Then
        Atualiza_Lista_Observacao_Pedido
   End If
   
   If chkCARREGAR_PEDIDO_BLOQUEIOS_PEDIDO.Value = vbChecked Then
        Atualiza_Lista_Pedido_Bloqueios
   End If
      
   If chkCARREGAR_NOTA_FISCAL_PEDIDO.Value = vbChecked Then
        Atualiza_Lista_Nota_Fiscal
   End If
   
   txtNUM_PEDIDO.Text = gstrCod
   txtCNPJ.Text = gstrCNPJ
   txtRAZAO_SOCIAL.Text = gstrRAZAO_SOCIAL
   
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   
   Set frmDetalhe_Pedido = Nothing
         
   GravaPosicaoList lvwItens_Pedido_Pedido
   GravaPosicaoList lvwObservacao_Pedido_Pedido
   GravaPosicaoList lvwPedido_Bloqueios_Pedido
   GravaPosicaoList lvwNota_Fiscal_Pedido
   
      
   FechaLista lvwItens_Pedido_Pedido
   FechaLista lvwObservacao_Pedido_Pedido
   FechaLista lvwPedido_Bloqueios_Pedido
   FechaLista lvwNota_Fiscal_Pedido
         
   GravaCheckbox chkCARREGAR_ITENS_PEDIDO_PEDIDO
   GravaCheckbox chkCARREGAR_OBSERVACAO_PEDIDO_PEDIDO
   GravaCheckbox chkCARREGAR_PEDIDO_BLOQUEIOS_PEDIDO
   GravaCheckbox chkCARREGAR_NOTA_FISCAL_PEDIDO
   
   FechaForm Me
   
End Sub

Private Sub lvwItens_Pedido_Pedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   Dim intColuna As Integer
   
   intColuna = ColumnHeader.Index - 1
   
   Select Case intColuna
   
        Case 14
            intColuna = 16
        Case 15
            intColuna = 17
            
   End Select
        
   If lvwItens_Pedido_Pedido.SortKey = intColuna Then

       If lvwItens_Pedido_Pedido.SortOrder = lvwAscending Then
           lvwItens_Pedido_Pedido.SortOrder = lvwDescending
       Else
           lvwItens_Pedido_Pedido.SortOrder = lvwAscending
       End If

   Else

       lvwItens_Pedido_Pedido.SortKey = intColuna
       lvwItens_Pedido_Pedido.SortOrder = lvwAscending

   End If

   lvwItens_Pedido_Pedido.Sorted = True

End Sub

Private Sub lvwNota_Fiscal_Pedido_DblClick()
    
    On Error GoTo TrataErro
    
    If lvwNota_Fiscal_Pedido.ListItems.Count = 0 Then Exit Sub
    With frmDetalhe_Nota_Fiscal
        .Codigo = gstrCod
        .CNPJ = gstrCNPJ
        .RAZAO_SOCIAL = gstrRAZAO_SOCIAL
        .COD_NOTA_FISCAL = lvwNota_Fiscal_Pedido.SelectedItem.Text
        .SERIE = lvwNota_Fiscal_Pedido.SelectedItem.SubItems(1)
        .Show 1
    End With

    Exit Sub

TrataErro:

    Unload Me

End Sub

Private Sub lvwObservacao_Pedido_Pedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   If lvwObservacao_Pedido_Pedido.SortKey = ColumnHeader.Index - 1 Then

       If lvwObservacao_Pedido_Pedido.SortOrder = lvwAscending Then
           lvwObservacao_Pedido_Pedido.SortOrder = lvwDescending
       Else
           lvwObservacao_Pedido_Pedido.SortOrder = lvwAscending
       End If

   Else

       lvwObservacao_Pedido_Pedido.SortKey = ColumnHeader.Index - 1
       lvwObservacao_Pedido_Pedido.SortOrder = lvwAscending

   End If

   lvwObservacao_Pedido_Pedido.Sorted = True

End Sub

Private Sub lvwPedido_Bloqueios_Pedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   Dim intColuna As Integer
   
   intColuna = ColumnHeader.Index - 1
   
   Select Case intColuna
   
        Case 3
            intColuna = 6
                   
   End Select
        
   If lvwPedido_Bloqueios_Pedido.SortKey = intColuna Then

       If lvwPedido_Bloqueios_Pedido.SortOrder = lvwAscending Then
           lvwPedido_Bloqueios_Pedido.SortOrder = lvwDescending
       Else
           lvwPedido_Bloqueios_Pedido.SortOrder = lvwAscending
       End If

   Else

       lvwPedido_Bloqueios_Pedido.SortKey = intColuna
       lvwPedido_Bloqueios_Pedido.SortOrder = lvwAscending

   End If

   lvwPedido_Bloqueios_Pedido.Sorted = True

End Sub


Private Sub lvwNota_Fiscal_Pedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   Dim intColuna As Integer
   
   intColuna = ColumnHeader.Index - 1
   
   Select Case intColuna
   
        Case 6
            intColuna = 23
        Case 7
            intColuna = 24
            
   End Select
        
   If lvwNota_Fiscal_Pedido.SortKey = intColuna Then

       If lvwNota_Fiscal_Pedido.SortOrder = lvwAscending Then
           lvwNota_Fiscal_Pedido.SortOrder = lvwDescending
       Else
           lvwNota_Fiscal_Pedido.SortOrder = lvwAscending
       End If

   Else

       lvwNota_Fiscal_Pedido.SortKey = intColuna
       lvwNota_Fiscal_Pedido.SortOrder = lvwAscending

   End If

   lvwNota_Fiscal_Pedido.Sorted = True

End Sub

Private Sub tabDetalhe_Click()

    
    fraITENS_PEDIDO.Visible = False
    fraOBSERVACAO_PEDIDO.Visible = False
    fraPEDIDO_BLOQUEIOS.Visible = False
    fraNOTA_FISCAL.Visible = False
        
    fraOBSERVACAO_PEDIDO.Left = fraITENS_PEDIDO.Left
    fraPEDIDO_BLOQUEIOS.Left = fraITENS_PEDIDO.Left
    fraNOTA_FISCAL.Left = fraITENS_PEDIDO.Left
    
    Select Case UCase(tabDetalhe.SelectedItem.Key)
        
        Case "ITENS_PEDIDO"
            fraITENS_PEDIDO.Visible = True
            
        Case "OBSERVACAO_PEDIDO"
            fraOBSERVACAO_PEDIDO.Visible = True
                    
        Case "PEDIDO_BLOQUEIOS"
            fraPEDIDO_BLOQUEIOS.Visible = True
        
        Case "NOTA_FISCAL"
            fraNOTA_FISCAL.Visible = True
              
    End Select
    
End Sub

Private Sub Atualiza_Lista_Itens_Pedido()
   
   On Error GoTo ValidaErro
   Me.MousePointer = vbHourglass
   
    Dim Rst As adodb.Recordset
   
    Dim itmX As ListItem
   
    
    Dim fldNUM_PEDIDO
    Dim fldCOD_PRODUTO
    Dim fldID_SEQUENCIAL
    Dim fldQTD_PEDIDA
    Dim fldQTD_FATURADA
    Dim fldQTD_DESTINADA
    Dim fldQTD_EMPENHADA
    Dim fldSITUACAO
    Dim fldVLR_PRECO
    Dim fldVLR_SALDO
    Dim fldVLR_UNITARIO
    Dim fldDESCONTO
    Dim fldCOND_PAGTO
    Dim fldTAB_PRECO
    Dim fldDES_PRODUTO_CURTA
    Dim fldDES_PRODUTO_LONGA
    Dim fldDATA_CEDO
    Dim fldDATA_TARDE
    Dim fldDATA_BASE
    Dim fldGRP_FISC_PRC
    Dim fldOPER_FISC_PRC
    Dim fldGRP_FISC_ENT
    Dim fldOPER_FISC_ENT
    
    Dim strORDENA_DATA As String
    
    Set Rst = New adodb.Recordset
      
    lvwItens_Pedido_Pedido.ListItems.Clear
        
    Set Rst = Listar_Itens_Pedido(gstrCod)
   
    Set fldNUM_PEDIDO = Rst.Fields("NUM_PEDIDO")
    Set fldCOD_PRODUTO = Rst.Fields("COD_PRODUTO")
    Set fldID_SEQUENCIAL = Rst.Fields("ID_SEQUENCIAL")
    Set fldQTD_PEDIDA = Rst.Fields("QTD_PEDIDA")
    Set fldQTD_FATURADA = Rst.Fields("QTD_FATURADA")
    Set fldQTD_DESTINADA = Rst.Fields("QTD_DESTINADA")
    Set fldQTD_EMPENHADA = Rst.Fields("QTD_EMPENHADA")
    Set fldSITUACAO = Rst.Fields("SITUACAO")
    Set fldVLR_PRECO = Rst.Fields("VLR_PRECO")
    Set fldVLR_SALDO = Rst.Fields("VLR_SALDO")
    Set fldVLR_UNITARIO = Rst.Fields("VLR_UNITARIO")
    Set fldDESCONTO = Rst.Fields("DESCONTO")
    Set fldCOND_PAGTO = Rst.Fields("COND_PAGTO")
    Set fldTAB_PRECO = Rst.Fields("TABELA_PRECO")
    Set fldDES_PRODUTO_CURTA = Rst.Fields("DES_PRODUTO_CURTA")
    Set fldDES_PRODUTO_LONGA = Rst.Fields("DES_PRODUTO_LONGA")
    Set fldDATA_CEDO = Rst.Fields("DATA_CEDO")
    Set fldDATA_TARDE = Rst.Fields("DATA_TARDE")
    Set fldDATA_BASE = Rst.Fields("DATA_BASE")
    Set fldGRP_FISC_PRC = Rst.Fields("GRP_FISCAL_PRC")
    Set fldOPER_FISC_PRC = Rst.Fields("OPER_FISCAL_PRC")
    Set fldGRP_FISC_ENT = Rst.Fields("GRP_FISCAL_ENT")
    Set fldOPER_FISC_ENT = Rst.Fields("OPER_FISCAL_ENT")
          
    If Rst.EOF Then
   
      With Me.lvwItens_Pedido_Pedido
          .ColumnHeaders.Clear
          .ListItems.Clear
          .ColumnHeaders.Add , , "Mensagem : N o existem registros selecionados.", 9000
      End With
   
    Else
   
      With lvwItens_Pedido_Pedido
          .ListItems.Clear
          With .ColumnHeaders
            .Clear
            .Add , , "Seq.", 500
            .Add , , "C d.Prod. Cliente", 500
            .Add , , "Descr. Curta Prod.", 500
            .Add , , "Descr. Longa Prod.", 500
            .Add , , "Qtde Pedida", 500, vbRightJustify
            .Add , , "Qtde Faturada", 500, vbRightJustify
            .Add , , "Qtde Destinada", 500, vbRightJustify
            .Add , , "Qtde Empenhada", 500, vbRightJustify
            .Add , , "Situa  o", 500, vbRightJustify
            .Add , , "Pre o", 500, vbRightJustify
            .Add , , "Saldo", 500, vbRightJustify
            .Add , , "Unit rio", 500, vbRightJustify
            .Add , , "Desconto", 500, vbRightJustify
            .Add , , "Condi  o Pagamento.", 1000
            .Add , , "Tabela Pre o", 500
            .Add , , "Data Base", 500
            .Add , , "Data Cedo", 500
            .Add , , "Data Tarde", 500
            .Add , , "Grupo Op.Fiscal", 500
            .Add , , "Oper. Fiscal", 500
            .Add , , "Grp Op.Fiscal Entrega", 500
            .Add , , "Oper. Fiscal Entrega", 500
            .Add , , "Data Cedo", 500
            .Add , , "Data Tarde", 500
         
         End With
      End With
      
      PreparaLista lvwItens_Pedido_Pedido
     
      With Rst.Fields
         
         Rst.MoveFirst
      
         Do While Not Rst.EOF
     
         
            Set itmX = lvwItens_Pedido_Pedido.ListItems.Add(, , fldID_SEQUENCIAL)
            
            itmX.SubItems(1) = IIf(Not Vazio(fldCOD_PRODUTO), fldCOD_PRODUTO, "")
            itmX.SubItems(2) = IIf(Not Vazio(fldDES_PRODUTO_CURTA), fldDES_PRODUTO_CURTA, "")
            itmX.SubItems(3) = IIf(Not Vazio(fldDES_PRODUTO_LONGA), fldDES_PRODUTO_LONGA, "")
            itmX.SubItems(4) = IIf(Not Vazio(Trim(fldQTD_PEDIDA)), ObterCampoNumerico(fldQTD_PEDIDA), "0")
            itmX.SubItems(5) = IIf(Not Vazio(Trim(fldQTD_FATURADA)), ObterCampoNumerico(fldQTD_FATURADA), "0")
            itmX.SubItems(6) = IIf(Not Vazio(Trim(fldQTD_DESTINADA)), ObterCampoNumerico(fldQTD_DESTINADA), "0")
            itmX.SubItems(7) = IIf(Not Vazio(Trim(fldQTD_EMPENHADA)), ObterCampoNumerico(fldQTD_EMPENHADA), "0")
            itmX.SubItems(8) = IIf(Not Vazio(Trim(fldSITUACAO)), fldSITUACAO, "")
            itmX.SubItems(9) = IIf(Not Vazio(Trim(fldVLR_PRECO)), ObterCampoNumerico(fldVLR_PRECO), "0")
            itmX.SubItems(10) = IIf(Not Vazio(Trim(fldVLR_SALDO)), ObterCampoNumerico(fldVLR_SALDO), "0")
            itmX.SubItems(11) = IIf(Not Vazio(Trim(fldVLR_UNITARIO)), ObterCampoNumerico(fldVLR_UNITARIO), "0")
            itmX.SubItems(12) = IIf(Not Vazio(Trim(fldDESCONTO)), ObterCampoNumerico(fldDESCONTO), "0")
            itmX.SubItems(13) = IIf(Not Vazio(Trim(fldCOND_PAGTO)), fldCOND_PAGTO, "")
            itmX.SubItems(14) = IIf(Not Vazio(Trim(fldTAB_PRECO)), fldTAB_PRECO, "")
            itmX.SubItems(15) = IIf(Not Vazio(Trim(fldDATA_BASE)), fldDATA_BASE, "")
            itmX.SubItems(16) = IIf(Not Vazio(Trim(fldDATA_CEDO)), fldDATA_CEDO, "")
            itmX.SubItems(17) = IIf(Not Vazio(Trim(fldDATA_TARDE)), fldDATA_TARDE, "")
            itmX.SubItems(18) = IIf(Not Vazio(fldGRP_FISC_PRC), fldGRP_FISC_PRC, "")
            itmX.SubItems(19) = IIf(Not Vazio(fldOPER_FISC_PRC), fldOPER_FISC_PRC, "")
            itmX.SubItems(20) = IIf(Not Vazio(fldGRP_FISC_ENT), fldGRP_FISC_ENT, "")
            itmX.SubItems(21) = IIf(Not Vazio(fldOPER_FISC_ENT), fldOPER_FISC_ENT, "")
                                                            
            If Not Vazio(fldDATA_CEDO) Then
                strORDENA_DATA = Right(fldDATA_CEDO, 4) & Mid(fldDATA_CEDO, 4, 2) & Left(fldDATA_CEDO, 2)
                itmX.SubItems(22) = strORDENA_DATA
            Else
                strORDENA_DATA = ""
                itmX.SubItems(22) = strORDENA_DATA
            End If
            
            If Not Vazio(fldDATA_TARDE) Then
                strORDENA_DATA = Right(fldDATA_TARDE, 4) & Mid(fldDATA_TARDE, 4, 2) & Left(fldDATA_TARDE, 2)
                itmX.SubItems(23) = strORDENA_DATA
            Else
                strORDENA_DATA = ""
                itmX.SubItems(23) = strORDENA_DATA
            End If
                                                            
            Rst.MoveNext
         Loop
      
      End With
      
      lvwItens_Pedido_Pedido.ColumnHeaders.Item(23).Width = 0
      lvwItens_Pedido_Pedido.ColumnHeaders.Item(24).Width = 0
   
    End If
   
    Dim intPosicao As Double
    
    intPosicao = RetornaPosicaoList(lvwItens_Pedido_Pedido)
   
    If intPosicao <> 0 Then
    
       lvwItens_Pedido_Pedido.ListItems.Item(intPosicao).Selected = True
    
    End If
   
    Set Rst = Nothing
   
    Me.MousePointer = vbDefault
   
    Exit Sub
ValidaErro:
   
   Me.MousePointer = vbDefault
   TrataErro Err.Number, Err.Description, Err.Source, True, Me.Caption
   
End Sub


Private Sub Atualiza_Lista_Observacao_Pedido()
   
   On Error GoTo ValidaErro
   Me.MousePointer = vbHourglass
   
   Dim Rst As adodb.Recordset
   
   Dim itmX As ListItem
      
   Dim fldNUM_PEDIDO
   Dim fldID_SEQUENCIAL
   Dim fldDES_TIPO_OPERACAO
   Dim fldINSCRICAO_ESTADUAL
   Dim fldNOME
   Dim fldENDERECO
   Dim fldMRH
   Dim fldCIDADE
   Dim fldUF
   Dim fldMUNICIPIO
   Dim fldTEXTO_NOTA_FISCAL
   Dim fldTEXTO_LIVRE
    
   Set Rst = New adodb.Recordset
      
   lvwObservacao_Pedido_Pedido.ListItems.Clear
        
   Set Rst = Listar_Observacao_Pedido(gstrCod)
   
   Set fldNUM_PEDIDO = Rst.Fields("NUM_PEDIDO")
   Set fldID_SEQUENCIAL = Rst.Fields("ID_SEQUENCIAL")
   Set fldDES_TIPO_OPERACAO = Rst.Fields("DES_TIPO_OPERACAO")
   Set fldINSCRICAO_ESTADUAL = Rst.Fields("INSCRICAO_ESTADUAL")
   Set fldNOME = Rst.Fields("NOME")
   Set fldENDERECO = Rst.Fields("ENDERECO")
   Set fldMRH = Rst.Fields("MRH")
   Set fldCIDADE = Rst.Fields("CIDADE")
   Set fldUF = Rst.Fields("UF")
   Set fldMUNICIPIO = Rst.Fields("MUNICIPIO")
   Set fldTEXTO_NOTA_FISCAL = Rst.Fields("TEXTO_NOTA_FISCAL")
   Set fldTEXTO_LIVRE = Rst.Fields("TEXTO_LIVRE")
          
   If Rst.EOF Then
   
      With Me.lvwObservacao_Pedido_Pedido
          .ColumnHeaders.Clear
          .ListItems.Clear
          .ColumnHeaders.Add , , "Mensagem : N o existem registros selecionados.", 9000
      End With
   
   Else
   
      With lvwObservacao_Pedido_Pedido
          .ListItems.Clear
          With .ColumnHeaders
            .Clear
            .Add , , "Seq.", 500
            .Add , , "Tipo Opera  o", 500
            .Add , , "Inscri  o Estadual Cliente", 500
            .Add , , "Nome", 500
            .Add , , "Endere o", 500
            .Add , , "Cidade", 500
            .Add , , "UF", 500
            .Add , , "Munic pio", 500
            .Add , , "MRH", 500
            .Add , , "Texto Nota Fiscal", 500
            .Add , , "Texto Livre", 500
         End With
      End With
      
      PreparaLista lvwObservacao_Pedido_Pedido
     
      With Rst.Fields
         
         Rst.MoveFirst
      
         Do While Not Rst.EOF
     
            Set itmX = lvwObservacao_Pedido_Pedido.ListItems.Add(, , fldID_SEQUENCIAL)
         
            itmX.SubItems(1) = IIf(Not Vazio(fldDES_TIPO_OPERACAO), fldDES_TIPO_OPERACAO, "")
            itmX.SubItems(2) = IIf(Not Vazio(fldINSCRICAO_ESTADUAL), fldINSCRICAO_ESTADUAL, "")
            itmX.SubItems(3) = IIf(Not Vazio(fldNOME), fldNOME, "")
            itmX.SubItems(4) = IIf(Not Vazio(fldENDERECO), fldENDERECO, "")
            itmX.SubItems(5) = IIf(Not Vazio(fldCIDADE), fldCIDADE, "")
            itmX.SubItems(6) = IIf(Not Vazio(fldUF), fldUF, "")
            itmX.SubItems(7) = IIf(Not Vazio(fldMUNICIPIO), fldMUNICIPIO, "")
            itmX.SubItems(8) = IIf(Not Vazio(fldMRH), fldMRH, "")
            itmX.SubItems(9) = IIf(Not Vazio(fldTEXTO_NOTA_FISCAL), fldTEXTO_NOTA_FISCAL, "")
            itmX.SubItems(10) = IIf(Not Vazio(fldTEXTO_LIVRE), fldTEXTO_LIVRE, "")
                                             
            Rst.MoveNext
         Loop
      
      End With
   
   End If
   
   Dim intPosicao As Double
    
   intPosicao = RetornaPosicaoList(lvwObservacao_Pedido_Pedido)
   
   If intPosicao <> 0 Then
    
       lvwObservacao_Pedido_Pedido.ListItems.Item(intPosicao).Selected = True
    
   End If
   
   Set Rst = Nothing
   
   Me.MousePointer = vbDefault
   
   Exit Sub
ValidaErro:
   
   Me.MousePointer = vbDefault
   TrataErro Err.Number, Err.Description, Err.Source, True, Me.Caption
   
End Sub



Private Sub Atualiza_Lista_Pedido_Bloqueios()
   
   On Error GoTo ValidaErro
   Me.MousePointer = vbHourglass
   
   Dim Rst As adodb.Recordset
   
   Dim itmX As ListItem
      
   Dim fldNUM_PEDIDO
   Dim fldNUM_LINHA
   Dim fldID_SEQUENCIAL
   Dim fldDES_BLOQUEIO
   Dim fldSTATUS
   Dim fldDATA_STATUS
   Dim fldDES_MENSAGEM
   Dim fldID_BLOQUEIO
      
   Dim strORDENA_DATA As String
    
   Set Rst = New adodb.Recordset
      
   lvwPedido_Bloqueios_Pedido.ListItems.Clear
        
   Set Rst = Listar_Pedido_Bloqueios(gstrCod)
   
   Set fldNUM_PEDIDO = Rst.Fields("NUM_PEDIDO")
   Set fldNUM_LINHA = Rst.Fields("NUM_LINHA")
   Set fldID_SEQUENCIAL = Rst.Fields("ID_SEQUENCIAL")
   Set fldDES_BLOQUEIO = Rst.Fields("DES_BLOQUEIO")
   Set fldSTATUS = Rst.Fields("STATUS")
   Set fldDATA_STATUS = Rst.Fields("DATA_STATUS")
   Set fldDES_MENSAGEM = Rst.Fields("DES_MENSAGEM")
   Set fldID_BLOQUEIO = Rst.Fields("ID_BLOQUEIO")
          
   If Rst.EOF Then
   
      With Me.lvwPedido_Bloqueios_Pedido
          .ColumnHeaders.Clear
          .ListItems.Clear
          .ColumnHeaders.Add , , "Mensagem : N o existem registros selecionados.", 9000
      End With
   
   Else
   
      With lvwPedido_Bloqueios_Pedido
          .ListItems.Clear
          With .ColumnHeaders
            .Clear
            .Add , , "Seq.", 500
            .Add , , "Descri  o Bloqueio", 500
            .Add , , "Status", 500
            .Add , , "Data do Status", 500
            .Add , , "Mensagem", 500
            .Add , , "Tipo Bloqueio", 500
            .Add , , "Data do Status", 500
            
         End With
      End With
      
      PreparaLista lvwPedido_Bloqueios_Pedido
     
      With Rst.Fields
         
         Rst.MoveFirst
      
         Do While Not Rst.EOF
     
         
            Set itmX = lvwPedido_Bloqueios_Pedido.ListItems.Add(, , IIf(Not Vazio(fldID_SEQUENCIAL), fldID_SEQUENCIAL, ""))
            
            itmX.SubItems(1) = IIf(Not Vazio(fldDES_BLOQUEIO), fldDES_BLOQUEIO, "")
            If fldSTATUS = "A" Then
                itmX.SubItems(2) = "Ativo"
            Else
                itmX.SubItems(2) = "Inativo"
            End If
            itmX.SubItems(3) = IIf(Not Vazio(fldDATA_STATUS), fldDATA_STATUS, "")
            itmX.SubItems(4) = IIf(Not Vazio(fldDES_MENSAGEM), fldDES_MENSAGEM, "")
                                                         
            Select Case fldID_BLOQUEIO
            
                Case "A"
                    itmX.SubItems(5) = "Altera  o"
                Case "B"
                    itmX.SubItems(5) = "Cobran a"
                Case "C"
                    itmX.SubItems(5) = "CLIENTE"
                Case "E"
                    itmX.SubItems(5) = "Frete"
                Case "F"
                    itmX.SubItems(5) = "Configurador"
                Case "G"
                    itmX.SubItems(5) = "Grupo Ordem"
                Case "H"
                    itmX.SubItems(5) = "Entrega"
                Case "K"
                    itmX.SubItems(5) = "S Comp Kit"
                Case "L"
                    itmX.SubItems(5) = "Local Entrega"
                Case "M"
                    itmX.SubItems(5) = "Margem Min / Max"
                Case "N"
                    itmX.SubItems(5) = "Pre o Min"
                Case "O"
                    itmX.SubItems(5) = "Vlr.Max Venda"
                Case "P"
                    itmX.SubItems(5) = "Produto"
                Case "Q"
                    itmX.SubItems(5) = "Qtde Min / Max"
                Case "R"
                    itmX.SubItems(5) = "Verif.Cr dito"
                Case "S"
                    itmX.SubItems(5) = "Venda"
                Case "T"
                    itmX.SubItems(5) = "Cotas"
                Case "U"
                    itmX.SubItems(5) = "Fech.Ciclo"
                Case "V"
                    itmX.SubItems(5) = "Vendor HLD"
                Case "X"
                    itmX.SubItems(5) = "Qtde Min/Max CO"
                Case "Y"
                    itmX.SubItems(5) = "Toler ncia.Zero"
                Case "Z"
                    itmX.SubItems(5) = "Prazo M dio"
            
            End Select
            
            If Not Vazio(fldDATA_STATUS) Then
                strORDENA_DATA = Right(fldDATA_STATUS, 4) & Mid(fldDATA_STATUS, 4, 2) & Left(fldDATA_STATUS, 2)
                itmX.SubItems(6) = strORDENA_DATA
            Else
                strORDENA_DATA = ""
                itmX.SubItems(6) = strORDENA_DATA
            End If
                                                                                                                              
            Rst.MoveNext
         Loop
      
      End With
                  
      lvwPedido_Bloqueios_Pedido.ColumnHeaders.Item(7).Width = 0
   
   End If
   
   Dim intPosicao As Double
    
   intPosicao = RetornaPosicaoList(lvwPedido_Bloqueios_Pedido)
   
   If intPosicao <> 0 Then
    
       lvwPedido_Bloqueios_Pedido.ListItems.Item(intPosicao).Selected = True
    
   End If
   
   Set Rst = Nothing
   
   Me.MousePointer = vbDefault
   
   Exit Sub
ValidaErro:
   
   Me.MousePointer = vbDefault
   TrataErro Err.Number, Err.Description, Err.Source, True, Me.Caption
   
End Sub

Private Sub Atualiza_Lista_Nota_Fiscal()
   
   On Error GoTo ValidaErro
   Me.MousePointer = vbHourglass
   
    Dim Rst As adodb.Recordset
   
    Dim itmX As ListItem
   
    
    Dim fldCOD_NOTA_FISCAL
    Dim fldSERIE
    Dim fldNUM_PEDIDO
    Dim fldCLIENTE
    Dim fldESTABELECIMENTO
    Dim fldCOD_FABRICA
    Dim fldSTATUS_NF
    Dim fldTIPO_NF
    Dim fldDATA_EMISSAO
    Dim fldDATA_SAIDA_MER
    Dim fldVALOR_BCICM
    Dim fldVALOR_ICM
    Dim fldVALOR_IPI
    Dim fldVALOR_ALIQICM
    Dim fldPESO_LIQ
    Dim fldPESO_BRUTO
    Dim fldVALOR_DESC
    Dim fldVALOR_TOTAL
    Dim fldTOTAL_UNID_FATUR
    Dim fldQTD_VOLUME
    Dim fldVIA_TRANSPORTE
    Dim fldDES_TRANSPORTE
    Dim fldVALOR_DESC_PONT
    Dim fldDES_QUALIDADE
    Dim fldCODMOEDA
    Dim fldDES_FABRICA
    
    Dim strORDENA_DATA As String
    
    Set Rst = New adodb.Recordset
      
    lvwNota_Fiscal_Pedido.ListItems.Clear
        
    Set Rst = Listar_Nota_Fiscal(, , gstrCod)
   
    Set fldCOD_NOTA_FISCAL = Rst.Fields("COD_NOTA_FISCAL")
    Set fldSERIE = Rst.Fields("SERIE")
    Set fldNUM_PEDIDO = Rst.Fields("NUM_PEDIDO")
    Set fldCLIENTE = Rst.Fields("CLIENTE")
    Set fldESTABELECIMENTO = Rst.Fields("ESTABELECIMENTO")
    Set fldCOD_FABRICA = Rst.Fields("COD_FABRICA")
    Set fldSTATUS_NF = Rst.Fields("STATUS_NF")
    Set fldTIPO_NF = Rst.Fields("TIPO_NF")
    Set fldDATA_EMISSAO = Rst.Fields("DATA_EMISSAO")
    Set fldDATA_SAIDA_MER = Rst.Fields("DATA_SAIDA_MER")
    Set fldVALOR_BCICM = Rst.Fields("VALOR_BCICM")
    Set fldVALOR_ICM = Rst.Fields("VALOR_ICM")
    Set fldVALOR_IPI = Rst.Fields("VALOR_IPI")
    Set fldVALOR_ALIQICM = Rst.Fields("VALOR_ALIQICM")
    Set fldPESO_LIQ = Rst.Fields("PESO_LIQ")
    Set fldPESO_BRUTO = Rst.Fields("PESO_BRUTO")
    Set fldVALOR_DESC = Rst.Fields("VALOR_DESC")
    Set fldVALOR_TOTAL = Rst.Fields("VALOR_TOTAL")
    Set fldTOTAL_UNID_FATUR = Rst.Fields("TOTAL_UNID_FATUR")
    Set fldQTD_VOLUME = Rst.Fields("QTD_VOLUME")
    Set fldVIA_TRANSPORTE = Rst.Fields("VIA_TRANSPORTE")
    Set fldDES_TRANSPORTE = Rst.Fields("DES_TRANSPORTE")
    Set fldVALOR_DESC_PONT = Rst.Fields("VALOR_DESC_PONT")
    Set fldDES_QUALIDADE = Rst.Fields("DES_QUALIDADE")
    Set fldCODMOEDA = Rst.Fields("CODMOEDA")
    Set fldDES_FABRICA = Rst.Fields("DES_FABRICA")
          
   If Rst.EOF Then
   
      With Me.lvwNota_Fiscal_Pedido
          .ColumnHeaders.Clear
          .ListItems.Clear
          .ColumnHeaders.Add , , "Mensagem : N o existem registros selecionados.", 9000
      End With
   
   Else
   
      With lvwNota_Fiscal_Pedido
          .ListItems.Clear
          With .ColumnHeaders
            .Clear
            .Add , , "C digo", 500
            .Add , , "S rie", 500, vbRightJustify
            .Add , , "Estabelecimento", 500, vbRightJustify
            .Add , , "F brica", 500
            .Add , , "Status NF", 500, vbRightJustify
            .Add , , "Tipo NF", 500, vbRightJustify
            .Add , , "Data Emiss o", 500, vbRightJustify
            .Add , , "Data Sa da Mercadoria", 500, vbRightJustify
            .Add , , "Valor BCICM", 500, vbRightJustify
            .Add , , "ICM", 500, vbRightJustify
            .Add , , "IPI", 500, vbRightJustify
            .Add , , "ALIQICM", 500, vbRightJustify
            .Add , , "Peso L quido", 500, vbRightJustify
            .Add , , "Peso Bruto", 500, vbRightJustify
            .Add , , "Valor Descr.", 500, vbRightJustify
            .Add , , "Valor Total", 500, vbRightJustify
            .Add , , "Total Unidade Faturada", 500, vbRightJustify
            .Add , , "Qtde Volume", 500, vbRightJustify
            .Add , , "VIA Transporte", 500, vbRightJustify
            .Add , , "Descr. Transporte", 500, vbRightJustify
            .Add , , "Valor Descr. Pont.", 500, vbRightJustify
            .Add , , "C d. Moeda", 500, vbRightJustify
            .Add , , "Qualidade", 500
            .Add , , "Data Emiss o", 500, vbRightJustify
            .Add , , "Data Sa da Mercadoria", 500, vbRightJustify
                        
         End With
      End With
      
      PreparaLista lvwNota_Fiscal_Pedido
     
      With Rst.Fields
         
         Rst.MoveFirst
      
         Do While Not Rst.EOF
     
        
     
         
            Set itmX = lvwNota_Fiscal_Pedido.ListItems.Add(, , fldCOD_NOTA_FISCAL)
            
            itmX.SubItems(1) = IIf(Not Vazio(fldSERIE), fldSERIE, "")
            itmX.SubItems(2) = IIf(Not Vazio(fldESTABELECIMENTO), fldESTABELECIMENTO, "")
            itmX.SubItems(3) = IIf(Not Vazio(fldDES_FABRICA), fldDES_FABRICA, "")
                        
            Select Case fldSTATUS_NF
              Case "A"
                itmX.SubItems(4) = "Contabilizado"
             Case "C"
                itmX.SubItems(4) = "Encerrado"
             Case "E"
                itmX.SubItems(4) = "Editado"
             Case "H"
                itmX.SubItems(4) = "Suspenso"
             Case "O"
                itmX.SubItems(4) = "Aberto"
             Case "V"
                itmX.SubItems(4) = "Transf.Voucher"
             Case "X"
                itmX.SubItems(4) = "Excluido"
            End Select
            
            itmX.SubItems(5) = IIf(Not Vazio(fldTIPO_NF), fldTIPO_NF, "")
            itmX.SubItems(6) = IIf(Not Vazio(fldDATA_EMISSAO), fldDATA_EMISSAO, "")
            itmX.SubItems(7) = IIf(Not Vazio(fldDATA_SAIDA_MER), fldDATA_SAIDA_MER, "")
            itmX.SubItems(8) = IIf(Not Vazio(Trim(fldVALOR_BCICM)), ObterCampoNumerico(fldVALOR_BCICM), "0")
            itmX.SubItems(9) = IIf(Not Vazio(Trim(fldVALOR_ICM)), ObterCampoNumerico(fldVALOR_ICM), "0")
            itmX.SubItems(10) = IIf(Not Vazio(Trim(fldVALOR_IPI)), ObterCampoNumerico(fldVALOR_IPI), "0")
            itmX.SubItems(11) = IIf(Not Vazio(Trim(fldVALOR_ALIQICM)), ObterCampoNumerico(fldVALOR_ALIQICM), "0")
            itmX.SubItems(12) = IIf(Not Vazio(Trim(fldPESO_LIQ)), ObterCampoNumerico(fldPESO_LIQ), "0")
            itmX.SubItems(13) = IIf(Not Vazio(Trim(fldPESO_BRUTO)), ObterCampoNumerico(fldPESO_BRUTO), "0")
            itmX.SubItems(14) = IIf(Not Vazio(Trim(fldVALOR_DESC)), ObterCampoNumerico(fldVALOR_DESC), "0")
            itmX.SubItems(15) = IIf(Not Vazio(Trim(fldVALOR_TOTAL)), ObterCampoNumerico(fldVALOR_TOTAL), "0")
            itmX.SubItems(16) = IIf(Not Vazio(Trim(fldTOTAL_UNID_FATUR)), ObterCampoNumerico(fldTOTAL_UNID_FATUR), "0")
            itmX.SubItems(17) = IIf(Not Vazio(Trim(fldQTD_VOLUME)), ObterCampoNumerico(fldQTD_VOLUME), "0")
            itmX.SubItems(18) = IIf(Not Vazio(fldVIA_TRANSPORTE), fldVIA_TRANSPORTE, "")
            itmX.SubItems(19) = IIf(Not Vazio(fldDES_TRANSPORTE), fldDES_TRANSPORTE, "")
            itmX.SubItems(20) = IIf(Not Vazio(Trim(fldVALOR_DESC_PONT)), ObterCampoNumerico(fldVALOR_DESC_PONT), "0")
            itmX.SubItems(21) = IIf(Not Vazio(fldCODMOEDA), fldCODMOEDA, "")
            itmX.SubItems(22) = IIf(Not Vazio(fldDES_QUALIDADE), fldDES_QUALIDADE, "")
                        
            If Not Vazio(fldDATA_EMISSAO) Then
                strORDENA_DATA = Right(fldDATA_EMISSAO, 4) & Mid(fldDATA_EMISSAO, 4, 2) & Left(fldDATA_EMISSAO, 2)
                itmX.SubItems(23) = strORDENA_DATA
            Else
                strORDENA_DATA = ""
                itmX.SubItems(23) = strORDENA_DATA
            End If
            
            If Not Vazio(fldDATA_SAIDA_MER) Then
                strORDENA_DATA = Right(fldDATA_SAIDA_MER, 4) & Mid(fldDATA_SAIDA_MER, 4, 2) & Left(fldDATA_SAIDA_MER, 2)
                itmX.SubItems(24) = strORDENA_DATA
            Else
                strORDENA_DATA = ""
                itmX.SubItems(24) = strORDENA_DATA
            End If
                                    
            Rst.MoveNext
         Loop
      
      End With
      
      lvwNota_Fiscal_Pedido.ColumnHeaders.Item(24).Width = 0
      lvwNota_Fiscal_Pedido.ColumnHeaders.Item(25).Width = 0
   
   End If
   
   Dim intPosicao As Double
    
   intPosicao = RetornaPosicaoList(lvwNota_Fiscal_Pedido)
   
   If intPosicao <> 0 Then
    
       lvwNota_Fiscal_Pedido.ListItems.Item(intPosicao).Selected = True
    
   End If
   
   Set Rst = Nothing
   
   Me.MousePointer = vbDefault
   
   Exit Sub
ValidaErro:
   
   Me.MousePointer = vbDefault
   TrataErro Err.Number, Err.Description, Err.Source, True, Me.Caption
   
End Sub

## CVGlobal.bas

Attribute VB_Name = "IFV01_GLOBAL"
Option Explicit

Public gstrDbms         As String
Public gstrUsuario      As String
Public gstrSenha        As String
Public gstrMembro       As String
Public gstrAdm_Sistema  As String
Public gstrVersaoWin    As String
Public gstrTpMembro     As String

Public Const gstrSubSistema         As String = "IFV"
Public Const gstrModuloSubSistema   As String = "01"


Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'--- Declare para IsTaskRunning
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103
Private Const STILL_ACTIVE = STATUS_PENDING

'------------------------------------------------------------------------------------
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion As String * 128
End Type
'-----------------------------------------------------------------------------------

Public Declare Function GetLocaleInfo Lib "kernel32" _
    Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
        ByVal lpLCData As String, ByVal cchData As Long) As Long
        
Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Const Locale_sthousand = &HF
Public Const Locale_sdecimal = &HE

Public symbolDECIMAL As String
Public iRet1DECIMAL As Long
Public iRet2DECIMAL As Long

Public symbolTHOUSAND As String
Public iRet1THOUSAND As Long
Public iRet2THOUSAND As Long

Public lpLCDatavar As String
Public Pos As Integer
Public Locale As Long

'
'  Descricao : Prepara Formulario
'  Retorno   : Posiciona Form
'

Public Function PreparaForm(ByVal frmFORM As Form)

    With frmFORM
        If .Tag = "" Then
            .Top = GetSetting("IFV01", gstrSubSistema, "Top" & .Caption, .Top)
            .Left = GetSetting("IFV01", gstrSubSistema, "Left" & .Caption, .Left)
            .WindowState = GetSetting("IFV01", gstrSubSistema, "WindowState " & .Caption, .WindowState)
        Else
            .Caption = LoadResString(.Tag)
            .Top = GetSetting("IFV01", gstrSubSistema, "Top" & .Tag, .Top)
            .Left = GetSetting("IFV01", gstrSubSistema, "Left" & .Tag, .Left)
            .WindowState = GetSetting("IFV01", gstrSubSistema, "WindowState " & .Tag, .WindowState)
        End If
        .Icon = LoadResPicture(101, vbResIcon)
    End With

End Function

'
'  Descricao : Fecha Formulario e grava posicao
'  Retorno   : -
'

Public Function FechaForm(frmFORM As Form)

    With frmFORM
        If .WindowState <> vbMinimized Then
            If .Tag = "" Then
                If .WindowState <> vbMaximized Then
                    SaveSetting "IFV01", gstrSubSistema, "Top" & .Caption, .Top
                    SaveSetting "IFV01", gstrSubSistema, "Left" & .Caption, .Left
                End If
                SaveSetting "IFV01", gstrSubSistema, "WindowState " & .Caption, .WindowState
            Else
                If .WindowState <> vbMaximized Then
                    SaveSetting "IFV01", gstrSubSistema, "Top" & .Tag, .Top
                    SaveSetting "IFV01", gstrSubSistema, "Left" & .Tag, .Left
                End If
                SaveSetting "IFV01", gstrSubSistema, "WindowState " & .Tag, .WindowState
            End If
        End If
    End With

End Function

'
'Descri  o:     TrataErro Generico
'Retorno    : Retorna Mensagem e Gera Log
'

Public Sub TrataErro(ByVal lngNumeroErro As Long, _
                     ByVal strDescricaoErro As String, _
                     ByVal strSource As String, _
                     ByVal blnShowUser As Boolean, _
                     ByVal strNOME_JANELA As String)

    If blnShowUser Then
    
       strSource = Replace(strSource, "->", Chr(13) & "->", 1)
        
       '  MsgBox CStr(lngNumeroErro) + " " + strDescricaoErro + " " + strSource, vbExclamation, App.EXEName
       MsgBox strDescricaoErro + Chr(13) + strSource, vbExclamation, strNOME_JANELA
              
    End If

End Sub

'
' Descricao : Prepara Main com opcao de seguranca por grupo usuario
' Retorno   : Formata Form Mdi
'

Private Sub PreparaFormMdi(ByVal frmFORM As Form)

    With frmFORM

        .StaStatusBar.Panels("Usuario").Text = UCase(gstrUsuario)
        ' .StaStatusBar.Panels("Banco").Text = UCase(mvarNomeDataSource)

        If .Tag <> "" Then
            .Caption = LoadResString(.Tag)
            .Top = GetSetting("IFV01", gstrSubSistema, "Top" & .Tag, .Top)
            .Left = GetSetting("IFV01", gstrSubSistema, "Left" & .Tag, .Left)
            .WindowState = GetSetting("IFV01", gstrSubSistema, "WindowState " & .Tag, .WindowState)
        End If
        '.Icon = LoadResPicture(101, vbResIcon)

    End With

End Sub

'
'  Descricao : Grava tamanho coluna do listview conforme op  o usu rio
'  Retorno   : -
'

Public Sub FechaLista(ByVal lvwLista As ListView)

    Dim intContador     As Integer
    Dim strListName     As String
    Dim intColunas      As Integer

    With lvwLista

        intColunas = .ColumnHeaders.Count
        strListName = .Name

        If intColunas > 1 Then
            For intContador = 1 To intColunas

                SaveSetting "IFV01", gstrSubSistema, strListName & "(" & CStr(intContador) & ")", .ColumnHeaders.Item(intContador).Width

            Next
        End If

    End With

End Sub

'
'  Descricao : Prepara ListView Formulario
'  Retorno   : Montar colunas do listview conforme op  o do usu rio
'
'
Public Sub PreparaLista(ByVal lvwLista As ListView, _
                        Optional ByVal cTamanho As Integer)

    Dim intContador     As Integer
    Dim strListName     As String
    Dim intColunas      As Integer
    
    If cTamanho = 0 Then
       cTamanho = 1440
    End If
    
    
    With lvwLista

        intColunas = .ColumnHeaders.Count
        strListName = .Name

        For intContador = 1 To intColunas
            If Replace(CStr(GetSetting("IFV01", gstrSubSistema, strListName & "(" & CStr(intContador) & ")", cTamanho)), ",", ".") <> "0" Then
                .ColumnHeaders.Item(intContador).Width = Val(Replace(CStr(GetSetting("IFV01", gstrSubSistema, strListName & "(" & CStr(intContador) & ")", cTamanho)), ",", "."))
            Else
                .ColumnHeaders.Item(intContador).Width = cTamanho
            End If
        Next

    End With

End Sub

Public Function GravaPosicaoList(lvwList As ListView)

    With lvwList
        If lvwList.ListItems.Count > 0 Then
            SaveSetting "IFV01", gstrSubSistema, lvwList.Name, lvwList.SelectedItem.Key
        Else
            SaveSetting "IFV01", gstrSubSistema, lvwList.Name, ""
        End If
    End With

End Function

Public Function RetornaPosicaoList(lvwList As ListView) As Double

    Dim strKey As String
    Dim intPosicao As Double
    Dim intCont  As Double

    With lvwList
        strKey = GetSetting("IFV01", gstrSubSistema, lvwList.Name, "")
        
        For intCont = 1 To lvwList.ListItems.Count
            If lvwList.ListItems(intCont).Key = strKey Then
                intPosicao = intCont
            End If
        Next
        
    End With

    RetornaPosicaoList = intPosicao
    
End Function


'
' Descri  o  : Remove acentua  o Caption Menu
' Retorno    : Retorna string formatada
'
Private Function RemoveCaracter(strNomeOpcao As String) As String

    Dim aryCaracter(13, 1)   As String
    Dim intContador         As Integer
    Dim intposicaoinicial   As Integer

    aryCaracter(0, 0) = " "
    aryCaracter(0, 1) = "A"

    aryCaracter(1, 0) = " "
    aryCaracter(1, 1) = "A"

    aryCaracter(2, 0) = " "
    aryCaracter(2, 1) = "A"

    aryCaracter(3, 0) = " "
    aryCaracter(3, 1) = "E"

    aryCaracter(4, 0) = " "
    aryCaracter(4, 1) = "E"

    aryCaracter(5, 0) = " "
    aryCaracter(5, 1) = "I"

    aryCaracter(6, 0) = " "
    aryCaracter(6, 1) = "I"

    aryCaracter(7, 0) = " "
    aryCaracter(7, 1) = "O"

    aryCaracter(8, 0) = " "
    aryCaracter(8, 1) = "O"

    aryCaracter(9, 0) = " "
    aryCaracter(9, 1) = "O"

    aryCaracter(10, 0) = " "
    aryCaracter(10, 1) = "U"

    aryCaracter(11, 0) = " "
    aryCaracter(11, 1) = "U"

    aryCaracter(12, 0) = " "
    aryCaracter(12, 1) = "C"

    aryCaracter(13, 0) = "&"
    aryCaracter(13, 1) = ""


    For intContador = 0 To 13

        intposicaoinicial = 0

        intposicaoinicial = InStr(strNomeOpcao, aryCaracter(intContador, 0))

        If intposicaoinicial > 0 Then

            strNomeOpcao = Mid$(strNomeOpcao, 1, intposicaoinicial - 1) & aryCaracter(intContador, 1) & Mid$(strNomeOpcao, intposicaoinicial + 1)

        End If

    Next

    RemoveCaracter = strNomeOpcao

End Function


'
'  Descricao  : Retorna Numero utilizado no Sql conforme o Dbms utilizado
'  Retorna    : --
'
Public Function GetNumero(ByVal dblNomeCampo As Double) As String

    GetNumero = Replace(CStr(dblNomeCampo), ",", ".")

End Function

'
'  Descricao  : Valida chamada sistema atraves do Menu de Sistemas
'  Retorna    : --
'

Public Function ValidaCommand()

    If InStr(UCase$(Command$), "/IFV01") <= 0 Or _
       InStr(UCase$(Command$), "/USER") <= 0 Then

       End

    End If

End Function

'recebe o valor em string e retorna o valor em fun  o
'do tipo string
Public Function fRetornaValor() As Boolean
    
    Locale = GetUserDefaultLCID()
    
    'Traz o separador de decimal
    iRet1DECIMAL = GetLocaleInfo(Locale, Locale_sdecimal, lpLCDatavar, 0)
    symbolDECIMAL = String$(iRet1DECIMAL, 0)
    iRet2DECIMAL = GetLocaleInfo(Locale, Locale_sdecimal, symbolDECIMAL, iRet1DECIMAL)
    
    'Traz o separador de milhares
    iRet1THOUSAND = GetLocaleInfo(Locale, Locale_sthousand, lpLCDatavar, 0)
    symbolTHOUSAND = String$(iRet1THOUSAND, 0)
    iRet2THOUSAND = GetLocaleInfo(Locale, Locale_sthousand, symbolTHOUSAND, iRet1THOUSAND)
    
    If InStr(1, symbolDECIMAL, ",") Then fRetornaValor = False
    If InStr(1, symbolDECIMAL, ".") Then fRetornaValor = True
        
    '--- Formato da Data "dd/MM/yyyy"
    Dim strDATA As String
    strDATA = DateSerial(2000, 1, 8)
    
    If strDATA <> "08/01/2000" Then
        fRetornaValor = False
    End If
    
End Function
    

'
' Descricao  : Verifica se Vazio ou Vazio
' Retorno    : Retorna true ou false
'

Public Function Vazio(vntVariavel As Variant) As Boolean

    Dim blnRetorno  As Boolean
    
    blnRetorno = True
    
    If Not IsMissing(vntVariavel) Then
        If Not IsNull(vntVariavel) Then
            If vntVariavel <> "" Then
                blnRetorno = False
            End If
        End If
    End If
    
    Vazio = blnRetorno

End Function

'
'  Descricao  : Formata a data utilizado no Sql conforme o Dbms utilizado
'  Retorna    : Retorna string sql
'
Public Function DataAtual() As String

    Dim strDataAtual As String
    
    strDataAtual = Format(Date, "DD/MM/YYYY")

    If gstrDbms = "ORACLE" Then
    
        DataAtual = "TO_DATE( '" & strDataAtual & "','DD/MM/YYYY')"
    
    Else
        
        DataAtual = "CAST( '" & Mid(strDataAtual, 1, 2) & "-" & NomeReduzidoMes(Val(Mid(strDataAtual, 4, 2))) & "-" & Mid(strDataAtual, 7, 4) & "' AS DATETIME )"
        
    End If
    
End Function

'
'  Descricao  : Retorna Data Formata a data utilizado no Sql conforme o Dbms utilizado
'  Retorna    : --
'
Public Function GetData(ByVal strNomeCampo As String) As String

    If gstrDbms = "ORACLE" Then
    
        GetData = "TO_CHAR(" & strNomeCampo & ",'DD/MM/YYYY')"
    
    Else
    
        GetData = "Convert(varchar(10)," & strNomeCampo & ", 103)"

    End If
    
End Function
'
'  Descricao  : Retorna Data Formata a data utilizado no Sql conforme o Dbms utilizado
'  Retorna    : --
'
Public Function GetHora(ByVal strNomeCampo As String) As String

    If gstrDbms = "ORACLE" Then
    
        GetHora = "" & strNomeCampo & ""
    
    Else
    
        GetHora = "Convert(varchar(8)," & strNomeCampo & ", 108)"

    End If
    
End Function
'
'  Descricao  : Formata a data utilizado no Sql conforme o Dbms utilizado
'  Retorna    : Retorna string sql
'
Public Function ToData(ByVal strDATA As String) As String

    If gstrDbms = "ORACLE" Then
    
        ToData = "TO_DATE( '" & strDATA & "','DD/MM/YYYY')"
    
    Else
       'LCOC
       ToData = "CONVERT(datetime, '" & Mid(strDATA, 7, 4) & "-" & Mid(strDATA, 4, 2) & "-" & Mid(strDATA, 1, 2) & "',21 )"
       'ToData = "CAST( '" & Mid(strDATA, 1, 2) & "-" & NomeReduzidoMes(Val(Mid(strDATA, 4, 2))) & "-" & Mid(strDATA, 7, 4) & "' AS DATETIME )"
       'If Mid(gstrVersaoWin, 1, 13) = "Windows Vista" Then
       '
       '   ToData = "CAST( '" & Mid(strDATA, 7, 4) & "-" & Mid(strDATA, 4, 2) & "-" & Mid(strDATA, 1, 2) & "' AS DATETIME )"
       'Else
       '   ToData = "CAST( '" & Mid(strDATA, 7, 4) & "-" & Mid(strDATA, 4, 2) & "-" & Mid(strDATA, 1, 2) & "' AS DATETIME )"
       'End If
    End If
    
    ToData = IIf(Vazio(strDATA), "NULL", ToData)
    
End Function
'
'  Descricao  : Formata a data utilizado no Sql conforme o Dbms utilizado
'  Retorna    : Retorna string sql
'
Public Function ToHora(ByVal strHora As String) As String

    If gstrDbms = "ORACLE" Then
    
        ToHora = "" & strHora & ""
    
    Else
    
        ToHora = "CAST( '" & strHora & "' AS DATETIME )"
        
    End If
    
    ToHora = IIf(Vazio(strHora), "NULL", ToHora)
    
End Function


'
' Descricao  : Recebe Lista de parametros separados por ponto(.) e
'              monta string Sql p/ ser utilizada na clausula In ( .. )
' Retorno    : Retorna string sql p/ utilizar na clausula In
'

Public Function MontaLista(ByVal strLista As String, intTamanhoString As Integer) As String

    
    Dim lngPosicao   As Long
    Dim lngTamanho   As Long
    Dim strString    As String
        
    lngPosicao = 1
    lngTamanho = Len(strLista)
    
    Do While lngPosicao < lngTamanho
    
        strString = IIf(Len(strString) > 0, strString & ",", "") & "'" & Mid(strLista, lngPosicao, intTamanhoString) & "'"
    
        lngPosicao = lngPosicao + intTamanhoString + 1
    
    Loop
    
    MontaLista = strString
    
End Function

'
'  Descricao  : Formata a data utilizado no Sql conforme o Dbms utilizado
'  Retorna    : Retorna string sql
'

Public Function GetConcat() As String

    If gstrDbms = "ORACLE" Then
    
        GetConcat = "||"
    
    Else
    
        GetConcat = "+"
        
    End If
    
End Function

'
' Descri  o : Substituir Caracter - Aspas Simples
' Retorno   : Retorna string formatado confome banco de dados utilizado
'

Public Function Substitui(ByVal strParametro As String) As String

    Dim strRetorno      As String
    
    strRetorno = Replace(strParametro, "'", IIf(gstrDbms = "ORACLE", "'||CHR(39)||'", "'+CHAR(39)+'"))
    
    Substitui = strRetorno
    
End Function

'
' Descri  o  : L  o register e Prepara String p/ Coneccao Banco de Dados
' Retorno    : String para conectar ao DBMS
'
'

Public Function ConnectionSQL(Optional ByVal strSECNAME As String) As String

    Dim strDbms                     As String
    Dim strDbmsUsuario              As String
    Dim strDbmsSenha                As String
    Dim strDataSource               As String
    Dim strInitialCatalog           As String
    Dim strConnect                  As String
    
    ConnectionSQL = ""
    
    Dim strFile As String
    
    strFile = App.Path & "\IFV.INI"
    
    If Vazio(Trim(strSECNAME)) Then
        strSECNAME = "IFV"
    End If
 
    If Len(Dir(strFile)) = 0 Then
      strDbmsUsuario = ""
      strDbmsSenha = ""
      strDataSource = ""
      strInitialCatalog = ""
    Else
      strDbmsUsuario = ReadWriteINI(strFile, "GET", strSECNAME, "Usuario")
      strDbmsSenha = ReadWriteINI(strFile, "GET", strSECNAME, "Senha")
      strDataSource = ReadWriteINI(strFile, "GET", strSECNAME, "DataSource")
      strInitialCatalog = ReadWriteINI(strFile, "GET", strSECNAME, "InitialCatalog")
    End If
    
    ' Alterado para Possibilitar v rios DBs - Escrit rios Representantes
    
    If strInitialCatalog = "DB_IFV" Then
       strInitialCatalog = "DB_" & gstrMembro
    End If
    
    '
    ' Variavel Global
    '
    strDbms = "SQLSERVER"
    
    gstrDbms = strDbms
    
    '
    ' Valida parametros da Connection ORACLE
    '
    
    If strDbms = "ORACLE" Then
    
        If Vazio(strDbms) Or Vazio(strDbmsUsuario) Or Vazio(strDataSource) Then
            Exit Function
        End If
    
    '
    ' Valida parametros da Connection SQLSERVER
    '
    
    ElseIf strDbms = "SQLSERVER" Then
    
        If Vazio(strDbms) Or Vazio(strDbmsUsuario) Or _
           Vazio(strDataSource) Or Vazio(strInitialCatalog) Then
            Exit Function
        End If
        
    End If
    
    '
    ' Prepara Connection
    '
    If strDbms = "ORACLE" Then
    
        strConnect = "Provider=MSDAORA.1;" & _
                     "Password=" & strDbmsSenha & ";" & _
                     "User ID=" & strDbmsUsuario & ";" & _
                     "Data Source=" & strDataSource & ";" & _
                     "Persist Security Info=True"
    
    ElseIf strDbms = "SQLSERVER" Then
    

'        strConnect = "Provider=SQLNCLI;" & _

        strConnect = "Provider=SQLOLEDB.1;" & _
                     IIf(Len(strDbmsSenha) > 0, "Password=" & strDbmsSenha & ";", "") & _
                     "UID=" & strDbmsUsuario & ";" & _
                     "Initial Catalog=" & strInitialCatalog & ";" & _
                     "Data Source=" & strDataSource & ";" & _
                     "Persist Security Info=" & IIf(Len(strDbmsSenha) > 0, "True", "False") & _
                     ";Connect Timeout =1000"
                     
    End If
    
    ConnectionSQL = strConnect
    
End Function

'
' Descricao : Obter nome reduzido mes em ingl s
' Retorno   : Nome reduzido
'
Private Function NomeReduzidoMes(ByVal intMes As Integer) As String

    Select Case intMes
    
        Case 1: NomeReduzidoMes = "JAN"
        Case 2: NomeReduzidoMes = "FEB"
        Case 3: NomeReduzidoMes = "MAR"
        Case 4: NomeReduzidoMes = "APR"
        Case 5: NomeReduzidoMes = "MAY"
        Case 6: NomeReduzidoMes = "JUN"
        Case 7: NomeReduzidoMes = "JUL"
        Case 8: NomeReduzidoMes = "AUG"
        Case 9: NomeReduzidoMes = "SEP"
        Case 10: NomeReduzidoMes = "OCT"
        Case 11: NomeReduzidoMes = "NOV"
        Case 12: NomeReduzidoMes = "DEC"
        
    End Select
    
End Function
'
' Descricao : Obter nome reduzido mes em ingl s
' Retorno   : Nome reduzido
'
Public Function NomeRedMesPort(ByVal intMes As Integer) As String

    Select Case intMes
    
        Case 1: NomeRedMesPort = "JAN"
        Case 2: NomeRedMesPort = "FEV"
        Case 3: NomeRedMesPort = "MAR"
        Case 4: NomeRedMesPort = "ABR"
        Case 5: NomeRedMesPort = "MAI"
        Case 6: NomeRedMesPort = "JUN"
        Case 7: NomeRedMesPort = "JUL"
        Case 8: NomeRedMesPort = "AGO"
        Case 9: NomeRedMesPort = "SET"
        Case 10: NomeRedMesPort = "OUT"
        Case 11: NomeRedMesPort = "NOV"
        Case 12: NomeRedMesPort = "DEZ"
        
    End Select
    
End Function

'
' Descricao : Obter o proximo numero de sequencia do registro especifico da tabela
' Retorno   : Retorna o proximo numero de sequencia de registro
'

Public Function ObterProximoRegistro(ByVal strSqlCampo As String, _
                                     ByVal strSQLFrom As String, _
                                     Optional ByVal strSQLWhere As Variant) As String

    Dim sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim strRetorno     As String

    strConnect = ConnectionSQL
    
    
    sql = "SELECT ISNULL( MAX(" & strSqlCampo & "), 0)" & strSqlCampo & " " & Chr(10)
    sql = sql & "FROM   " & strSQLFrom & " " & Chr(10)
    
    If Not Vazio(strSQLWhere) Then

        sql = sql & "WHERE " & strSQLWhere & " " & Chr(10)

    End If
    
    Set Rst = New adodb.Recordset
    
    Rst.MaxRecords = 1
    Rst.CursorLocation = adUseClient
    Rst.Open sql, strConnect, adOpenForwardOnly, adLockReadOnly
    
    strRetorno = "1"
    
    If Not Rst.EOF Then
        
        strRetorno = IIf(Not Vazio(Rst.Fields(0)), Rst.Fields(0) + 1, strRetorno)
        
    End If
    
    '
    ' Formatar Retorno atraves do  blnFormataZeros , blnQtdeFormataZeros
    '

    ObterProximoRegistro = strRetorno

End Function


'
' Descricao : Obter o proximo numero de sequencia do registro especifico da tabela
' Retorno   : Retorna o proximo numero de sequencia de registro
'

Public Function ReadWriteINI(FileName As String, Mode As String, tmpSecname As String, tmpKeyname As String, Optional tmpKeyValue) As String
  
  Dim tmpString As String
  Dim secname As String
  Dim KeyName As String
  Dim keyvalue As String
  Dim anInt
  Dim defaultkey As String
  Dim CurP As String

  On Error GoTo ReadWriteINIError
  '
  ' *** set the return value to OK
  'ReadWriteINI = "OK"
  ' *** test for good data to work with
  If IsNull(Mode) Or Len(Mode) = 0 Then
    ReadWriteINI = "ERROR MODE"    ' Set the return value
    Exit Function
  End If
  If IsNull(tmpSecname) Or Len(tmpSecname) = 0 Then
    ReadWriteINI = "ERROR Secname" ' Set the return value
    Exit Function
  End If
  If IsNull(tmpKeyname) Or Len(tmpKeyname) = 0 Then
    ReadWriteINI = "ERROR Keyname" ' Set the return value
    Exit Function
  End If
  ' *** set the ini file name
  '
  ' ******* WRITE MODE *************************************
  If UCase(Mode) = "WRITE" Then
    If IsNull(tmpKeyValue) Or Len(tmpKeyValue) = 0 Then
      ReadWriteINI = "ERROR KeyValue"
      Exit Function
    Else
      secname = tmpSecname
      KeyName = tmpKeyname
      keyvalue = tmpKeyValue
      anInt = WritePrivateProfileString(secname, KeyName, keyvalue, FileName)
    End If
  End If
  ' *******************************************************
  '
  ' *******  READ MODE *************************************
  If UCase(Mode) = "GET" Then
    secname = tmpSecname
    KeyName = tmpKeyname
    defaultkey = "Failed"
    keyvalue = String$(50, 32)
    anInt = GetPrivateProfileString(secname, KeyName, defaultkey, keyvalue, Len(keyvalue), FileName)
    If Left(keyvalue, 6) <> "Failed" Then        ' *** got it
      tmpString = keyvalue
      tmpString = RTrim(tmpString)
      tmpString = Left(tmpString, Len(tmpString) - 1)
    End If
    ReadWriteINI = tmpString
  End If
Exit Function
   
  ' *******
ReadWriteINIError:
   MsgBox Error
   Stop
End Function

Public Function codifica(strVALOR As String) As String
    
    'fun  o que criptografa
    Dim strResultado As String
    Dim chave        As String
    Dim intContador As Integer, intSoma As Integer
    
    chave = Chr(10) & Chr(1) & Chr(9) & Chr(2) & Chr(8) & Chr(3) & Chr(7) & Chr(4) & Chr(6) & Chr(5)
    strResultado = ""
    
    For intContador = 1 To Len(strVALOR)
        intSoma = intSoma + 1
        If intSoma = 11 Then intSoma = 1
        strResultado = strResultado & Chr$((Asc(Mid(strVALOR, intContador, 1)) - Asc(Mid(chave, intSoma, 1))))
    Next intContador
    
    codifica = strResultado
End Function

Public Function decodifica(strVALOR As String) As String
    'fun  o que descriptografa a senha do usu rio quando ele digita
    Dim strResultado As String
    Dim chave        As String
    Dim intContador As Integer, intSoma As Integer
    
    chave = Chr(10) & Chr(1) & Chr(9) & Chr(2) & Chr(8) & Chr(3) & Chr(7) & Chr(4) & Chr(6) & Chr(5)
    strResultado = ""
    
    For intContador = 1 To Len(strVALOR)
        intSoma = intSoma + 1
        If intSoma = 11 Then intSoma = 1
        strResultado = strResultado & Chr((Asc(Mid(strVALOR, intContador, 1)) + Asc(Mid(chave, intSoma, 1))))
    Next intContador
    
    decodifica = strResultado
    
End Function

'
' Descri  o  : Retorna o dominio
' Retorno    : String para conectar ao Dom nio
'
'
Public Function Retorna_Dominio() As String

    Dim strDominio    As String
        
    Dim strFile As String
    
    strFile = App.Path & "\IFV.INI"
    
    If Len(Dir(strFile)) > 0 Then
        strDominio = ReadWriteINI(strFile, "GET", "WINDOWS", "Dominio")
    End If
        
    Retorna_Dominio = strDominio
    
End Function

'
' Descri  o  : Retorna o dominio
' Retorno    : String para conectar ao Dom nio
'
'
Public Function Retorna_Grupo_Dominio() As String

    Dim strGrupo_Dominio    As String
        
    Dim strFile As String
    
    strFile = App.Path & "\IFV.INI"
    
    If Len(Dir(strFile)) > 0 Then
        strGrupo_Dominio = ReadWriteINI(strFile, "GET", "WINDOWS", "Grupo")
    End If
        
    Retorna_Grupo_Dominio = strGrupo_Dominio
    
End Function

Public Function SetErrSource(strNomeModulo As String, strNomeProcesso As String) As String

    SetErrSource = Chr(10) & Chr(10) & strNomeModulo & "->" & strNomeModulo & "." & strNomeProcesso & "@" & GetComputerName()
    
End Function

'
' Descri  o  : Obter o nome do computador ( DCOM application. )
' Retorno    : Retorna string com o nome do computador
'
Function GetComputerName() As String

    Dim strBuffer As String
    Dim lngLen As Long
        
    strBuffer = Space(255 + 1)
    lngLen = Len(strBuffer)
    
    If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
        GetComputerName = Left$(strBuffer, lngLen)
    Else
        GetComputerName = ""
    End If
    
End Function

'
' Descricao : Seleciona conteudo do TextBox
' Retorno   : --
'
Public Function SelecionaTexto(ByRef txtObject As TextBox)

    txtObject.SelStart = 0
    txtObject.SelLength = Len(txtObject.Text)

End Function

Public Function ObterCampoNumerico(ByVal strString As String, _
                                   Optional ByVal intDecimal As Integer = 2) As String

    Dim intContX As Integer

    Dim strDecimal As String
    
    For intContX = 1 To intDecimal
        strDecimal = strDecimal & "0"
    Next

    If intDecimal <> 0 Then
        strString = Format(strString, "###,###,###,###,###,###,###,###,###,##0." & strDecimal)
    Else
        strString = Format(strString, "###,###,###,###,###,###,###,###,###,##0")
    End If
    
    ObterCampoNumerico = Replace(strString, Chr(0), ",")
    
End Function


Public Function GravaCheckbox(strCheckbox As CheckBox)

    With strCheckbox
        SaveSetting "SSU2001", gstrSubSistema, strCheckbox.Name, strCheckbox.Value
    End With

End Function

Public Function RetornaCheckbox(strCheckbox As CheckBox)
         
    strCheckbox.Value = GetSetting("SSU2001", gstrSubSistema, strCheckbox.Name, vbUnchecked)
    
End Function


Public Function GerarExcel_ListView(lvwListView As ListView)


    If lvwListView.ListItems.Count <= 0 Then
        Exit Function
    End If

    '----[ Abrindo a Aplica  o ]
    Dim xCelFim, TPlanilha As String
    Dim Planilha As Excel.Application
    Set Planilha = New Excel.Application

    With Planilha
        '---[ Configuracoes ]
        .Visible = True
        .Workbooks.Add
        .ActiveWindow.Zoom = 100
        .ActiveSheet.PageSetup.Orientation = xlLandscape
        .DisplayAlerts = False '--- [ Nao exibe Alertas ]
        .ActiveWindow.View = xlNormalView
        
        '---[ Cabe alho ]
        Dim intColunas  As Integer
        Dim intLinhas   As Integer
        Dim intCont     As Integer
        Dim IntCont2    As Integer
        Dim intContador As Integer
        Dim Xyz As String
        
        intColunas = lvwListView.ColumnHeaders.Count - 1
        intLinhas = lvwListView.ListItems.Count
        intContador = 0
        
        For intCont = 0 To intColunas
            
            If lvwListView.ColumnHeaders(intCont + 1).Width <> 0 Then
                
                .Range(RetornaColunaExcel(intContador) & "1").Value = lvwListView.ColumnHeaders(intCont + 1).Text
                .Range(RetornaColunaExcel(intContador) & "1").Font.Bold = True
            
                For IntCont2 = 1 To intLinhas
                    If intCont = 0 Then
                        .Range(RetornaColunaExcel(intContador) & IntCont2 + 1).Value = lvwListView.ListItems.Item(IntCont2).Text
                    Else
                        Xyz = lvwListView.ListItems.Item(IntCont2).SubItems(intCont)
                        If Mid(Xyz, 3, 1) = "/" And Mid(Xyz, 6, 1) = "/" Then
                          .Range(RetornaColunaExcel(intContador) & IntCont2 + 1).Value = Mid(Xyz, 7, 4) & "/" & Mid(Xyz, 4, 2) & "/" & Mid(Xyz, 1, 2)
                        Else
                          .Range(RetornaColunaExcel(intContador) & IntCont2 + 1).Value = lvwListView.ListItems.Item(IntCont2).SubItems(intCont)
                        End If
                    End If
                Next
                
                intContador = intContador + 1
                
            End If
            xCelFim = RetornaColunaExcel(intContador - 1) & IntCont2
            
        Next

        '--- [ Ajustando Colunas ]
        .Columns("A:" & RetornaColunaExcel(intColunas)).Select
        .Selection.Columns.AutoFit
        
        .Range("A1").Select

    End With

    TPlanilha = Su_ColocaBorda("A1", xCelFim, Planilha)
    Set Planilha = Nothing

End Function

Public Function RetornaColunaExcel(ByVal intCont As Integer) As String

    If intCont < 25 Then
        RetornaColunaExcel = Chr(Asc("A") + intCont)
    End If

    If intCont >= 25 And intCont <= 50 Then
        RetornaColunaExcel = "A" & Chr(Asc("A") + intCont - 25)
    End If
    
    If intCont >= 51 And intCont <= 76 Then
        RetornaColunaExcel = "B" & Chr(Asc("A") + intCont - 51)
    End If
    
    If intCont >= 77 And intCont <= 102 Then
        RetornaColunaExcel = "C" & Chr(Asc("A") + intCont - 77)
    End If
    
    If intCont >= 103 And intCont <= 128 Then
        RetornaColunaExcel = "D" & Chr(Asc("A") + intCont - 103)
    End If
    
    If intCont >= 129 And intCont <= 154 Then
        RetornaColunaExcel = "E" & Chr(Asc("A") + intCont - 129)
    End If
    
    If intCont >= 155 And intCont <= 180 Then
        RetornaColunaExcel = "F" & Chr(Asc("A") + intCont - 155)
    End If
    
    If intCont >= 181 And intCont <= 206 Then
        RetornaColunaExcel = "G" & Chr(Asc("A") + intCont - 181)
    End If
    
    If intCont >= 207 And intCont <= 232 Then
        RetornaColunaExcel = "H" & Chr(Asc("A") + intCont - 207)
    End If
    
    If intCont >= 233 And intCont <= 254 Then
        RetornaColunaExcel = "I" & Chr(Asc("A") + intCont - 233)
    End If
    

End Function

Public Function Su_ColocaBorda(ByVal strCelulaInicio As String, _
                          ByVal strCelulaTermino As String, _
                          ByVal Excel_Planilha As Excel.Application)
    
    Excel_Planilha.Range(strCelulaInicio & ":" & strCelulaTermino).Select
    With Excel_Planilha
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        If strCelulaInicio <> strCelulaTermino Then
            With .Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
    End With
End Function

Function RetornaFormatadoCNPJCPF(ByVal strRetorno As String) As String

    If Not IsNull(strRetorno) Then

    If Len(Trim(strRetorno)) = 11 Then
        strRetorno = Mid(strRetorno, 1, 3) & "." & Mid(strRetorno, 4, 3) & "." & Mid(strRetorno, 7, 3) & "-" & Mid(strRetorno, 10)
    Else
        strRetorno = Mid(strRetorno, 1, 2) & "." & Mid(strRetorno, 3, 3) & "." & Mid(strRetorno, 6, 3) & "/" & Mid(strRetorno, 9, 4) & "-" & Mid(strRetorno, 13)
    End If

    End If
    RetornaFormatadoCNPJCPF = strRetorno

End Function

Function RetornaFormatadoIE_CEP(ByVal strRetorno As String) As String
    If Not IsNull(strRetorno) Then
    
       If InStr(1, strRetorno, ".") = 0 Then
    
          If Len(Trim(strRetorno)) = 8 Then
             strRetorno = Mid(strRetorno, 1, 5) & "-" & Mid(strRetorno, 6, 3)
          Else
             If Len(Trim(strRetorno)) >= 10 Then
                strRetorno = Mid(strRetorno, 1, 3) & "." & Mid(strRetorno, 4, 3) & "." & Mid(strRetorno, 7, 3) & "." & Mid(strRetorno, 10)
             End If
          End If
       End If
    End If
    RetornaFormatadoIE_CEP = strRetorno
End Function

Function RetornaFormatadoFONE(ByVal strRetorno As String) As String
    
    If Not IsNull(strRetorno) Then

    If Len(Trim(strRetorno)) >= 10 Then
       strRetorno = "(" & Mid(Trim(strRetorno), 1, 2) & ") " & Mid(strRetorno, 3)
    End If
    
    End If
    RetornaFormatadoFONE = strRetorno

End Function


Public Sub CarregaComboSituacao_Pedido(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_Situacao_Pedido()
    
    objCombo.ListField = "DES_SIT_PEDIDO"
    objCombo.BoundColumn = "COD_SIT_PEDIDO"
    
    Exit Sub
    
ValidaErro:
    TrataErro Err.Number, Err.Description, Err.Source, True, ""
End Sub

Public Sub CarregaComboSetor(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_RI_Setor()
    
    objCombo.ListField = "DES_SETOR"
    objCombo.BoundColumn = "COD_SETOR"
    
    Exit Sub
    
ValidaErro:
    TrataErro Err.Number, Err.Description, Err.Source, True, ""
End Sub

Public Sub CarregaComboLinha_Produto(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_Linha()
    
    objCombo.ListField = "LINHA_PRODUTO"
    objCombo.BoundColumn = "LINHA_PRODUTO"
    
    Exit Sub
    
ValidaErro:
    TrataErro Err.Number, Err.Description, Err.Source, True, ""
End Sub
Public Sub CarregaComboFamiliaCota(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_Familia_Cota()
    
    objCombo.ListField = "FAMILIA"
    objCombo.BoundColumn = "FAMILIA"
    
    Exit Sub
    
ValidaErro:
    TrataErro Err.Number, Err.Description, Err.Source, True, ""
End Sub

Public Sub CarregaComboCateg_Produto(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_Categoria()
    
    objCombo.ListField = "CATEGORIA_PROD"
    objCombo.BoundColumn = "CATEGORIA_PROD"
    
    Exit Sub
    
ValidaErro:
    TrataErro Err.Number, Err.Description, Err.Source, True, ""
End Sub

Public Sub CarregaComboData(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_Data()
    
    objCombo.ListField = "DATA_REQUERIDA"
    objCombo.BoundColumn = "DATA_REQUERIDA"
    
    Exit Sub
    
ValidaErro:
    TrataErro Err.Number, Err.Description, Err.Source, True, ""
End Sub


Public Sub CarregaComboCliente(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    Set objCombo.RowSource = Listar_Cliente()
    
    objCombo.ListField = "RAZAO_SOCIAL"
    objCombo.BoundColumn = "CLIENTE"
    
    Exit Sub
    
ValidaErro:

    TrataErro Err.Number, Err.Description, Err.Source, True, ""
    
End Sub

Public Sub CarregaComboTp_Estoque(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    objCombo.AddItem ""
    objCombo.ItemData(objCombo.NewIndex) = 0
    objCombo.AddItem "A - Amostra"
    objCombo.ItemData(objCombo.NewIndex) = 1
    objCombo.AddItem "F - Fora"
    objCombo.ItemData(objCombo.NewIndex) = 2
    objCombo.AddItem "H - Retalho"
    objCombo.ItemData(objCombo.NewIndex) = 3
    objCombo.AddItem "N - Normal"
    objCombo.ItemData(objCombo.NewIndex) = 4
    objCombo.AddItem "P - Ponta"
    objCombo.ItemData(objCombo.NewIndex) = 5
    objCombo.AddItem "R - R. Nuance"
    objCombo.ItemData(objCombo.NewIndex) = 6
    objCombo.AddItem "S - Slow"
    objCombo.ItemData(objCombo.NewIndex) = 7
    
    Exit Sub
    
ValidaErro:

    TrataErro Err.Number, Err.Description, Err.Source, True, ""
    
End Sub
Public Sub CarregaComboTp_Bloqueio(objCombo As Object)
                             
    On Error GoTo ValidaErro
    
    objCombo.AddItem ""
    objCombo.ItemData(objCombo.NewIndex) = 0
    objCombo.AddItem "C - Comercial"
    objCombo.ItemData(objCombo.NewIndex) = 1
    objCombo.AddItem "F - Financeiro"
    objCombo.ItemData(objCombo.NewIndex) = 2
    objCombo.AddItem "P - Planejamento"
    objCombo.ItemData(objCombo.NewIndex) = 3
    
    Exit Sub
    
ValidaErro:

    TrataErro Err.Number, Err.Description, Err.Source, True, ""
    
End Sub




'
'Descri  o  : Chama o Relat rio
'Retorno    : Chama o relat rio para execu  o fazendo o tratamento de poss veis erros
'

Public Sub Chama_Relatorio_Geral(ByVal rptCrystal As Object, _
                                 ByVal strReport As String, _
                                 Optional ByVal strSelectionFormula As String, _
                                 Optional ByVal strFormula0 As String, _
                                 Optional ByVal strFormula1 As String, _
                                 Optional ByVal strFormula2 As String, _
                                 Optional ByVal strFormula3 As String, _
                                 Optional ByVal strFormula4 As String, _
                                 Optional ByVal strFormula5 As String, _
                                 Optional ByVal strFormula6 As String, _
                                 Optional ByVal strFormula7 As String, _
                                 Optional ByVal strFormula8 As String)

    On Error GoTo ValidaErro

    Dim strPath As String
    Dim strModuloSubSistema As String
    Dim intContX As Integer

    strPath = UCase(App.Path)

    strModuloSubSistema = "\" & gstrSubSistema & gstrModuloSubSistema

    If strPath Like "*\FONT\CLIENT" Then

        strPath = Mid(strPath, 1, InStr(1, strPath, "\FONT\CLIENT") - 1)

    End If

    If strPath Like "*\EXE" Then

        strPath = Mid(strPath, 1, InStr(1, strPath, "\EXE") - 1)

    End If
        
        strPath = strPath & "\Report\" & strReport

    If Dir(strPath) = "" Then

            MsgBox "Arquivo n o encontrado - " & strPath, vbCritical, "Relat rio Crystal Report"

    Else

        rptCrystal.ReportFileName = strPath
        
        Dim strConnect As String
        
        strConnect = PreparaConeccaoReport
        'strConnect = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=SA;Initial Catalog=SALEXMARK;Data Source=CONVERGE_NT_03"

        rptCrystal.Connect = strConnect
                
        rptCrystal.DataFiles(0) = strConnect
        
        If Not Vazio(strSelectionFormula) Then

            rptCrystal.SelectionFormula = strSelectionFormula
            
        Else
            
            rptCrystal.SelectionFormula = ""
            
        End If
        
        'Zerar as formulas para n o evitar problemas com chamadas de outros reports
        
        For intContX = 0 To 8
            rptCrystal.Formulas(intContX) = ""
        Next
        
        If Not Vazio(strFormula0) Then
            rptCrystal.Formulas(0) = strFormula0
        End If
        
        If Not Vazio(strFormula1) Then
            rptCrystal.Formulas(1) = strFormula1
        End If
        
        If Not Vazio(strFormula2) Then
            rptCrystal.Formulas(2) = strFormula2
        End If
        
        If Not Vazio(strFormula3) Then
            rptCrystal.Formulas(3) = strFormula3
        End If

        If Not Vazio(strFormula4) Then
            rptCrystal.Formulas(4) = strFormula4
        End If
        
        If Not Vazio(strFormula5) Then
            rptCrystal.Formulas(5) = strFormula5
        End If

        If Not Vazio(strFormula6) Then
            rptCrystal.Formulas(6) = strFormula6
        End If
        
        If Not Vazio(strFormula7) Then
            rptCrystal.Formulas(7) = strFormula7
        End If

        If Not Vazio(strFormula8) Then
            rptCrystal.Formulas(8) = strFormula8
        End If
        rptCrystal.WindowState = crptMaximized
        rptCrystal.Action = 1

    End If

    Exit Sub

ValidaErro:

      TrataErro Err.Number, Err.Description, Err.Source, True, ""
      
      

End Sub

Public Function PreparaConeccaoReport() As String
    
    Dim strConnect As String
    
    On Error GoTo ErrorHandler
        
    strConnect = ConnectionReport("IFV")
    
    PreparaConeccaoReport = strConnect
    
    Exit Function
    
ErrorHandler:
    If Err.Number <> 0 Then
        TrataErro Err.Number, Err.Description, Err.Source, True, ""
    End If
    
End Function


'
' Descri  o  : L  o register e Prepara String p/ Coneccao Banco de Dados
' Retorno    : String para conectar ao DBMS
'
'

Public Function ConnectionReport(ByVal strSECNAME As String, _
                                 Optional strDbms As String = "SQLSERVER") As String

    Dim strDbmsUsuario              As String
    Dim strDbmsSenha                As String
    Dim strDataSource               As String
    Dim strInitialCatalog           As String
    Dim strConnect                  As String
            
    ConnectionReport = ""
    
    Dim strFile As String
    
    strFile = App.Path & "\IFV.INI"
 
    If Len(Dir(strFile)) = 0 Then
      strDbmsUsuario = ""
      strDbmsSenha = ""
      strDataSource = ""
      strInitialCatalog = ""
    Else
      strDbmsUsuario = ReadWriteINI(strFile, "GET", strSECNAME, "Usuario")
      strDbmsSenha = ReadWriteINI(strFile, "GET", strSECNAME, "Senha")
      strDataSource = ReadWriteINI(strFile, "GET", strSECNAME, "DataSource")
      strInitialCatalog = ReadWriteINI(strFile, "GET", strSECNAME, "InitialCatalog")
    End If
    
    ' Alterado para Possibilitar v rios DBs - Escrit rios Representantes
    
    If strInitialCatalog = "DB_IFV" Then
       strInitialCatalog = "DB_" & gstrMembro
    End If
    
    '
    ' Variavel Global
    '
    
    gstrDbms = strDbms
    
    '
    ' Valida parametros da Connection ORACLE
    '
    
    If strDbms = "ORACLE" Then
    
        If Vazio(strDbms) Or Vazio(strDbmsUsuario) Or Vazio(strDataSource) Then
    
            Exit Function
        
        End If
    
    '
    ' Valida parametros da Connection SQLSERVER
    '
    
    ElseIf strDbms = "SQLSERVER" Then
    
        If Vazio(strDbms) Or Vazio(strDbmsUsuario) Or _
           Vazio(strDataSource) Or Vazio(strInitialCatalog) Then
            Exit Function
        
        End If
        
    End If
    
    '
    ' Prepara Connection
    '
    If strDbms = "ORACLE" Then
    
        strConnect = "Provider=MSDAORA.1;" & _
                     "PWD=" & strDbmsSenha & ";" & _
                     "UID=" & strDbmsUsuario & ";" & _
                     "Data Source=" & strDataSource

        
        
    ElseIf strDbms = "SQLSERVER" Then
                     
       strConnect = "Driver={SQL SERVER};" & _
                     "Server=" & strDataSource & ";" & _
                     "UID=" & strDbmsUsuario & ";" & _
                     IIf(Len(strDbmsSenha) > 0, "PWD=" & strDbmsSenha & ";", "") & _
                     "Database=" & strInitialCatalog
                     
    End If
    
    ConnectionReport = strConnect
    
End Function

Public Function RetornaBanco(ByVal strSECNAME As String) As String

    Dim strInitialCatalog               As String
            
    RetornaBanco = ""
    
    Dim strFile As String
    
    strFile = App.Path & "\IFV.INI"
 
    If Len(Dir(strFile)) = 0 Then
      strInitialCatalog = ""
    Else
      strInitialCatalog = ReadWriteINI(strFile, "GET", strSECNAME, "InitialCatalog")
    End If
        
    If Vazio(Trim(strInitialCatalog)) Then
        Exit Function
    End If
        
    ' Alterado para Possibilitar v rios DBs - Escrit rios Representantes
    
    If strInitialCatalog = "DB_IFV" Then
       strInitialCatalog = "DB_" & gstrMembro
    End If
        
    RetornaBanco = Trim(strInitialCatalog)
    
End Function

Public Function RetornaVersao(ByVal strSECNAME As String) As String

    Dim strVersao               As String
            
    RetornaVersao = ""
    
    Dim strFile As String
    
    strFile = App.Path & "\IFV.INI"
 
    If Len(Dir(strFile)) = 0 Then
      strVersao = ""
    Else
      strVersao = ReadWriteINI(strFile, "GET", strSECNAME, "Versao")
    End If
        
    If Vazio(Trim(strVersao)) Then
        Exit Function
    End If
        
    RetornaVersao = Trim(strVersao)
    
End Function

Public Function IsTaskRunning(ByVal hApp As Long) As Boolean
    
    Dim hProc As Long
    Dim lExitCode As Long
    
    IsTaskRunning = False
    
    On Error GoTo noslm:
    
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hApp)
    GetExitCodeProcess hProc, lExitCode
    If lExitCode = STILL_ACTIVE Then
        IsTaskRunning = True
    End If

noslm:

End Function

'
' Descric o : Buscar_Membro
' Retorno   : RecordSet
'

Public Function Buscar_Membro(Optional ByVal strCOD_MEMBRO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    sql = " SELECT " & Chr(10)
    sql = sql & "   COD_MEMBRO," & Chr(10)
    sql = sql & "   NOME_MEMBRO," & Chr(10)
    sql = sql & "   SITUACAO," & Chr(10)
    sql = sql & "   SENHA," & Chr(10)
    sql = sql & "   EMAIL," & Chr(10)
    sql = sql & "   TIPO_MEMBRO" & Chr(10)


    sql = sql & "   FROM MEMBROS " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_MEMBRO)) Then
         sql = sql & SqlAux & " COD_MEMBRO = '" & Substitui(strCOD_MEMBRO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    sql = sql & "ORDER BY COD_MEMBRO" & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Buscar_Membro = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource("Login", "Buscar_Membro"), Err.Description

End Function

'
' Descric o : Montar_Lista_Flds
' Retorno   : RecordSet
'
Public Function Montar_Listar_Flds(Optional ByVal strNomeModulo As String) As adodb.Recordset

    On Error GoTo ErrorHandler
    
    Dim sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    Dim Tabelas_V(10) As String
    Dim Contar         As Integer
    
    If Vazio(strNomeModulo) Then
       strNomeModulo = "IFV"
    End If

    strConnect = ConnectionSQL("IFV")

    sql = " SELECT DISTINCT TABELAS FROM LISTA_TAB " & Chr(10)
        
    sql = sql & "ORDER BY TABELAS " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open sql, strConnect, adOpenForwardOnly, adLockReadOnly
    
    If Rst.EOF Then
   
       Set Montar_Listar_Flds = Rst
   
   Else
      
       Contar = 1
       
       Rst.MoveFirst
         
       Do While Not Rst.EOF
         
          Tabelas_V(Contar) = Rst.Fields("TABELAS")
            
          Contar = Contar + 1
            
          Rst.MoveNext
       Loop
      
   End If

   Set Montar_Listar_Flds = Rst

   Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Montar_Listar_Flds"), Err.Description

End Function

Public Function VersaoWin(Optional ExibirDetalhes As Boolean = False) As String
 
 Dim Estrutura As OSVERSIONINFO, Ver As String, Build As String, ServicePack As String
 
 Estrutura.dwOSVersionInfoSize = Len(Estrutura)
 
 GetVersionEx Estrutura
 
 With Estrutura
 
        Build = .dwBuildNumber
        If Not Build = "" Then Build = " Build " & Build
        ServicePack = .szCSDVersion
        If Not ServicePack = "" Then ServicePack = " " & ServicePack
 
        If ExibirDetalhes Then
           Ver = ", vers o " & Format$(.dwMajorVersion, "0") & "." & Format$(.dwMinorVersion, "00") & Build & ServicePack
        Else
           Ver = ""
        End If
 
        Select Case .dwPlatformId
        Case 0, 1
            Select Case .dwMinorVersion
            Case 0
                VersaoWin = "Windows95" & Ver
            Case 10
                VersaoWin = "Windows98" & Ver
            Case Else
                VersaoWin = "Windows" & Ver
            End Select
        Case 2
            If InStr(1, Ver, "vers o 5") > 0 Then
               VersaoWin = "Windows XP" & Ver
            Else
               If InStr(1, Ver, "vers o 6") > 0 Then
                  VersaoWin = "Windows Vista" & Ver
               Else
                  VersaoWin = "Windows NT (NT4, 2000)" & Ver
               End If
            End If
        Case Else
            VersaoWin = "Sistema operacional n o identificado"
        End Select
 
 End With

End Function


Function RemoverAcessos() As Boolean
' Funcao criada para remover toda a configuracao quando a data dos dados fica maior que X dias
    Screen.MousePointer = vbHourglass

    Dim strDATA_CARGA As String
    Dim nDias As Integer
    Dim sFilename As String
    
    strDATA_CARGA = frmMain.BuscarDt_UltAtualizacao()
    'strDATA_CARGA = "07/06/2023"
    'nDias = DateDiff("d", Date, Format$(strDATA_CARGA, "dd/mm/yyyy"))
    nDias = DateDiff("d", strDATA_CARGA, Date)
  
    If nDias > 3 Then
    
        MsgBox "Ambiente desatualizado" + vbCrLf + "Baixar novamente os dados", vbOKOnly
        
        'RemoveArquivo "C:\IFV\EXE\REGISTRO.DAT"
        'RemoveArquivo "C:\IFV\DATA\CLIENTE.TXT"
        'fazer rotina para matar toda a pasta DATA menos arquivos MDF e LDF em C:\IFV\DATA
        
    
        sFilename = Dir("C:\IFV\DATA\") 'Dir(App.Path & "\Forms\")
        Do While sFilename > ""
           'If (Right(sFilename, 4) <> ".mdf") And (Right(sFilename, 4) <> ".ldf") Then
           If (UCase(sFilename) <> UCase("registro.dat")) And (UCase(sFilename) <> UCase("registro.txt")) _
            And (Right(UCase(sFilename), 4) <> ".MDF") And (Right(UCase(sFilename), 4) <> ".LDF") Then
                RemoveArquivo "C:\IFV\DATA\" & sFilename
           End If
           sFilename = Dir("C:\IFV\DATA\")
        Loop
        
        Screen.MousePointer = vbDefault
        End
    End If
    
   
    Screen.MousePointer = vbDefault
    
End Function

Function RemoveArquivo(cArquivo As String)
    On Error Resume Next
    Dim InFile As Integer
    If Dir(cArquivo, vbArchive) <> "" Then
       Kill cArquivo
       Open cArquivo For Output As InFile
       Print #InFile, "nulo"
       Close InFile
       Kill cArquivo
    End If
End Function


Public Sub Truncate_Tabelas()

    On Error GoTo ErrorHandler
        
    Dim strConnect     As String
    Dim sql            As String
    Dim cnn            As adodb.Connection
    
   ' sql = "DELETE FROM ITENS_NOTA_FISCAL" & Chr(10)
   ' sql = sql & "DELETE FROM DUPLICATAS" & Chr(10)
   ' sql = sql & "DELETE FROM NOTA_FISCAL" & Chr(10)
    
   ' sql = sql & "DELETE FROM ITENS_PEDIDO" & Chr(10)
   ' sql = sql & "DELETE FROM OBSERVACAO_PEDIDO" & Chr(10)
   ' sql = sql & "DELETE FROM PEDIDO_BLOQUEIOS" & Chr(10)
   ' sql = sql & "DELETE FROM PEDIDO" & Chr(10)
    
   ' sql = sql & "DELETE FROM ESTOQUE_PRODUTO" & Chr(10)
   ' sql = sql & "DELETE FROM PRODUTO" & Chr(10)
   ' sql = sql & "DELETE FROM PREVISAO_COTA_VENDA" & Chr(10)
   ' sql = sql & "DELETE FROM COR" & Chr(10)
   ' sql = sql & "DELETE FROM DIMENSAO" & Chr(10)
   ' sql = sql & "DELETE FROM EMBALAGEM" & Chr(10)
   ' sql = sql & "DELETE FROM ARTIGO_PADRAO" & Chr(10)
    
        
   ' sql = sql & "DELETE FROM EQUIPES_CLIENTE" & Chr(10)
   ' sql = sql & "DELETE FROM MEMBROS_EQUIPES" & Chr(10)
   ' sql = sql & "DELETE FROM MEMBROS" & Chr(10)
   ' sql = sql & "DELETE FROM EQUIPES" & Chr(10)
    
   ' sql = sql & "DELETE FROM RI_DETALHES" & Chr(10)
   ' sql = sql & "DELETE FROM RI_REGISTRO_INSATISFACAO" & Chr(10)
   ' sql = sql & "DELETE FROM RI_SETORES" & Chr(10)
        
   ' sql = sql & "DELETE FROM ATIVIDADES_CLIENTE" & Chr(10)
   ' sql = sql & "DELETE FROM CONTATO_CLIENTE" & Chr(10)
   ' sql = sql & "DELETE FROM CREDITO_CLIENTE" & Chr(10)
   ' sql = sql & "DELETE FROM CLIENTE" & Chr(10)
    
   ' sql = sql & "DELETE FROM FABRICA" & Chr(10)
    
   ' sql = sql & "DELETE FROM PRODUTO_CONCORRENTE" & Chr(10)
   ' sql = sql & "DELETE FROM FORNECEDOR" & Chr(10)
    
   ' sql = sql & "DELETE FROM TMP_CLIENTE" & Chr(10)
   ' sql = sql & "DELETE FROM TMP_PEDIDO" & Chr(10)
   ' sql = sql & "DELETE FROM TMP_PRODUTO" & Chr(10)

   '----- TRUNCATE TABLE  -------------------------
     sql = "TRUNCATE TABLE ITENS_NOTA_FISCAL" & Chr(10)
     sql = sql & "TRUNCATE TABLE DUPLICATAS" & Chr(10)
     sql = sql & "TRUNCATE TABLE NOTA_FISCAL" & Chr(10)
    
     sql = sql & "TRUNCATE TABLE ITENS_PEDIDO" & Chr(10)
     sql = sql & "TRUNCATE TABLE OBSERVACAO_PEDIDO" & Chr(10)
     sql = sql & "TRUNCATE TABLE PEDIDO_BLOQUEIOS" & Chr(10)
     sql = sql & "TRUNCATE TABLE PEDIDO" & Chr(10)
    
     sql = sql & "TRUNCATE TABLE ESTOQUE_PRODUTO" & Chr(10)
     sql = sql & "TRUNCATE TABLE PRODUTO" & Chr(10)
     sql = sql & "TRUNCATE TABLE PREVISAO_COTA_VENDA" & Chr(10)
     sql = sql & "TRUNCATE TABLE COR" & Chr(10)
     sql = sql & "TRUNCATE TABLE DIMENSAO" & Chr(10)
     sql = sql & "TRUNCATE TABLE EMBALAGEM" & Chr(10)
     sql = sql & "TRUNCATE TABLE ARTIGO_PADRAO" & Chr(10)
    
        
     sql = sql & "TRUNCATE TABLE EQUIPES_CLIENTE" & Chr(10)
     sql = sql & "TRUNCATE TABLE MEMBROS_EQUIPES" & Chr(10)
     sql = sql & "TRUNCATE TABLE MEMBROS" & Chr(10)
     sql = sql & "TRUNCATE TABLE EQUIPES" & Chr(10)
    
     sql = sql & "TRUNCATE TABLE RI_DETALHES" & Chr(10)
     sql = sql & "TRUNCATE TABLE RI_REGISTRO_INSATISFACAO" & Chr(10)
     sql = sql & "TRUNCATE TABLE RI_SETORES" & Chr(10)
        
     sql = sql & "TRUNCATE TABLE ATIVIDADES_CLIENTE" & Chr(10)
     sql = sql & "TRUNCATE TABLE CONTATO_CLIENTE" & Chr(10)
     sql = sql & "TRUNCATE TABLE CREDITO_CLIENTE" & Chr(10)
     sql = sql & "TRUNCATE TABLE CLIENTE" & Chr(10)
    
     sql = sql & "TRUNCATE TABLE FABRICA" & Chr(10)
    
     sql = sql & "TRUNCATE TABLE PRODUTO_CONCORRENTE " & Chr(10)
     sql = sql & "TRUNCATE TABLE FORNECEDOR " & Chr(10)
     '-- TB NOVO --
     sql = sql & "TRUNCATE TABLE ND " & Chr(10)
    
     sql = sql & "TRUNCATE TABLE TMP_CLIENTE" & Chr(10)
     sql = sql & "TRUNCATE TABLE TMP_PEDIDO" & Chr(10)
     sql = sql & "TRUNCATE TABLE TMP_PRODUTO" & Chr(10)


    strConnect = ConnectionSQL("IFV")
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute sql, , adExecuteNoRecords
        .Close
    End With
    
    Exit Sub

ErrorHandler:

    If Not cnn Is Nothing Then
        cnn.Close
        Set cnn = Nothing
    End If

    Err.Raise Err.Number, SetErrSource("Limpar Arquivo", "Truncate_Tabelas"), Err.Description

End Sub



## IFV01_CLIENTE.bas
Attribute VB_Name = "IFV01_CLIENTE"
Private Const strNomeModulo = "IFV01_CLIENTE"

'
' Descric o : Listar_Cliente
' Retorno   : RecordSet
'

Public Function Listar_Cliente(Optional ByVal strCLIENTE As String, _
                               Optional ByVal strCNPJ As String, _
                               Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                               Optional ByVal strUF As String, _
                               Optional ByVal strLIKE_CIDADE As String, _
                               Optional ByVal strStatus As String, _
                               Optional ByVal strNIVEL_CLIENTE As String, _
                               Optional ByVal strATENDIDO_CLI As String, _
                               Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT DISTINCT " & Chr(10)
    Sql = Sql & "   CLIENTE.CLIENTE , " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ , " & Chr(10)
    Sql = Sql & "   CLIENTE.ESTABELECIMENTO , " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL , " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA , " & Chr(10)
    Sql = Sql & "   CLIENTE.DATA_CADASTRAMENTO , " & Chr(10)
    Sql = Sql & "   CLIENTE.LOGRADOURO , " & Chr(10)
    Sql = Sql & "   CLIENTE.NUM_LOGRADOURO , " & Chr(10)
    Sql = Sql & "   CLIENTE.COMPL_LOGRADOURO , " & Chr(10)
    Sql = Sql & "   CLIENTE.BAIRRO , " & Chr(10)
    Sql = Sql & "   CLIENTE.CIDADE , " & Chr(10)
    Sql = Sql & "   CLIENTE.UF , " & Chr(10)
    Sql = Sql & "   CLIENTE.CEP , " & Chr(10)
    Sql = Sql & "   CLIENTE.NIVEL_CLIENTE , " & Chr(10)
    Sql = Sql & "   CLIENTE.INSCRICAO_ESTADUAL , " & Chr(10)
    Sql = Sql & "   CONTATOS.NUM_FONE, " & Chr(10)
    Sql = Sql & "   CONTATOS.NUM_FAX, " & Chr(10)
    Sql = Sql & "   CONTATOS.NOME_CONTATO, " & Chr(10)
    Sql = Sql & "   CLIENTE.COBRANCA , " & Chr(10)
    Sql = Sql & "   CLIENTE.ENTREGA , " & Chr(10)
    Sql = Sql & "   CLIENTE.COMERCIAL,  " & Chr(10)
    Sql = Sql & "   CLIENTE.STATUS, " & Chr(10)
    Sql = Sql & "   CLIENTE.ATENDIDO " & Chr(10)
    
    Sql = Sql & "   FROM CLIENTE CLIENTE " & Chr(10)
    
    Sql = Sql & " LEFT JOIN CONTATO_CLIENTE CONTATOS" & Chr(10)
    Sql = Sql & " ON(CLIENTE.CLIENTE = CONTATOS.CLIENTE" & Chr(10)
    Sql = Sql & " AND CONTATOS.ID_SEQUENCIAL = 1)" & Chr(10)

    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & " INNER JOIN ATIVIDADES_CLIENTE ATIV_CLIENTE" & Chr(10)
         Sql = Sql & " ON(CLIENTE.CLIENTE = ATIV_CLIENTE.CLIENTE " & Chr(10)
         Sql = Sql & " AND ATIV_CLIENTE.UNIDADE_NEGOCIO = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
    End If

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CLIENTE.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCNPJ)) Then
         Sql = Sql & SqlAux & " CLIENTE.CNPJ = '" & Substitui(strCNPJ) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & " (CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strUF)) Then
         Sql = Sql & SqlAux & " CLIENTE.UF = '" & Substitui(strUF) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_CIDADE)) Then
         Sql = Sql & SqlAux & " CLIENTE.CIDADE LIKE '%" & Substitui(strLIKE_CIDADE) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
       
    If Not Vazio(Trim(strStatus)) Then
         Sql = Sql & SqlAux & " CLIENTE.STATUS = '" & Substitui(strStatus) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strNIVEL_CLIENTE)) Then
         Sql = Sql & SqlAux & " CLIENTE.NIVEL_CLIENTE = '" & Substitui(strNIVEL_CLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strATENDIDO_CLI)) Then
       If strATENDIDO_CLI <> "A" Then
         Sql = Sql & SqlAux & " CLIENTE.ATENDIDO = '" & Substitui(strATENDIDO_CLI) & "'" & Chr(10)
         SqlAux = " AND "
       End If
    End If
    
    Sql = Sql & "ORDER BY CLIENTE.RAZAO_SOCIAL " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Cliente = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Cliente"), Err.Description

End Function

'
' Descric o : Listar_Contato_Cliente
' Retorno   : RecordSet
'

Public Function Listar_Contato_Cliente(Optional ByVal strCLIENTE As String, _
                                       Optional ByVal strID_SEQUENCIAL As String, _
                                       Optional ByVal strLIKE_NOME_CONTATO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
        
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   CONT_CLIENTE.CLIENTE , " & Chr(10)
    Sql = Sql & "   CONT_CLIENTE.ID_SEQUENCIAL , " & Chr(10)
    Sql = Sql & "   CONT_CLIENTE.NOME_CONTATO , " & Chr(10)
    Sql = Sql & "   CONT_CLIENTE.NUM_FONE , " & Chr(10)
    Sql = Sql & "   CONT_CLIENTE.NUM_FAX , " & Chr(10)
    Sql = Sql & "   CONT_CLIENTE.EMAIL_CNT , " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)
    
    Sql = Sql & "   FROM CONTATO_CLIENTE CONT_CLIENTE " & Chr(10)
    Sql = Sql & "   INNER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "       ON(CLIENTE.CLIENTE = CONT_CLIENTE.CLIENTE) " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CONT_CLIENTE.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strID_SEQUENCIAL)) Then
         Sql = Sql & SqlAux & " CONT_CLIENTE.ID_SEQUENCIAL = " & strID_SEQUENCIAL & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_NOME_CONTATO)) Then
         Sql = Sql & SqlAux & " CONT_CLIENTE.NOME_CONTATO LIKE '%" & Substitui(strLIKE_NOME_CONTATO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    Sql = Sql & "ORDER BY CONT_CLIENTE.NOME_CONTATO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Contato_Cliente = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Contato_Cliente"), Err.Description

End Function

'
' Descric o : Listar_Atividades_Cliente
' Retorno   : RecordSet
'

Public Function Listar_Atividades_Cliente(Optional ByVal strCLIENTE As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
        
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.CLIENTE , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.UNIDADE_NEGOCIO , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.LINHA_PRODUTO , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.CANAL , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.RAMO , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.SUBRAMO , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.CATEGORIA , " & Chr(10)
    Sql = Sql & "   ATIV_CLIENTE.SITUACAO , " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)
    
    Sql = Sql & "   FROM ATIVIDADES_CLIENTE ATIV_CLIENTE " & Chr(10)
    Sql = Sql & "   INNER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "       ON(CLIENTE.CLIENTE = ATIV_CLIENTE.CLIENTE) " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " ATIV_CLIENTE.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY ATIV_CLIENTE.RAMO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Atividades_Cliente = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Atividades_Cliente"), Err.Description

End Function

'
' Descric o : Listar_Credito_Cliente
' Retorno   : RecordSet
'

Public Function Listar_Credito_Cliente(Optional ByVal strCLIENTE As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
        
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   CRED_CLIENTE.CLIENTE , " & Chr(10)
    Sql = Sql & "   CRED_CLIENTE.DESCRICAO , " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR,'0') VALOR, " & Chr(10)
    Sql = Sql & GetData("CRED_CLIENTE.DATA_RENOV_RESPALDO") & " DATA_RENOV_RESPALDO , " & Chr(10)
    Sql = Sql & GetData("CRED_CLIENTE.DATA_ULT_BALANCO") & " DATA_ULT_BALANCO , " & Chr(10)
    Sql = Sql & GetData("CRED_CLIENTE.DATA_ATRIBUICAO") & " DATA_ATRIBUICAO , " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_LIM_CREDITO,'0') VALOR_LIM_CREDITO, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_PATRIMONIO,'0') VALOR_PATRIMONIO, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_ACUMULO,'0') VALOR_ACUMULO, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_DIVIDA_VENCIDA,'0') VALOR_DIVIDA_VENCIDA, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_DIVIDA_VENCER,'0') VALOR_DIVIDA_VENCER, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_PEDLIBNFAT,'0') VALOR_PEDLIBNFAT, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_RETIDO,'0') VALOR_RETIDO, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_TIT_TOTAL,'0') VALOR_TIT_TOTAL, " & Chr(10)
    Sql = Sql & "   ISNULL(CRED_CLIENTE.VALOR_VENDOR_OPEN,'0') VALOR_VENDOR_OPEN, " & Chr(10)
    Sql = Sql & "   CRED_CLIENTE.DES_TRATAMENTO_DUPLIC , " & Chr(10)
    Sql = Sql & "   CRED_CLIENTE.DES_CONCEITO_BALANCO , " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)
    
    Sql = Sql & "   FROM CREDITO_CLIENTE CRED_CLIENTE " & Chr(10)
    Sql = Sql & "   INNER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "       ON(CLIENTE.CLIENTE = CRED_CLIENTE.CLIENTE) " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CRED_CLIENTE.CLIENTE LIKE '" & Mid(Substitui(strCLIENTE), 1, 8) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY CRED_CLIENTE.DESCRICAO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Credito_Cliente = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Credito_Cliente"), Err.Description

End Function


'
' Descric o : Listar_RI_Registro_Insatisfacao
' Retorno   : RecordSet
'

Public Function Listar_RI_Registro_Insatisfacao(Optional ByVal strNUM_RI As String, _
                                                Optional ByVal strLIKE_NUM_RI As String, _
                                                Optional ByVal strDATA_INCLUSAO As String, _
                                                Optional ByVal strDATA_JULGAMENTO As String, _
                                                Optional ByVal strCOD_SETOR As String, _
                                                Optional ByVal strCLIENTE As String, _
                                                Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                                                Optional ByVal strSTATUS_RI As String, _
                                                Optional ByVal strDATA_INCLUSAO_INI As String, _
                                                Optional ByVal strDATA_INCLUSAO_FIM As String, _
                                                Optional ByVal strDATA_JULGAMENTO_INI As String, _
                                                Optional ByVal strDATA_JULGAMENTO_FIM As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    If Not Vazio(Trim(strDATA_INCLUSAO_INI)) And Not Vazio(Trim(strDATA_INCLUSAO_FIM)) Then
    
        If DateValue(strDATA_INCLUSAO_INI) > DateValue(strDATA_INCLUSAO_FIM) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_RI_Registro_Insatisfacao"), "Data Fim Inclus o deve conter valor maior ou igual Data In cio Inclus o."
        End If
        
    End If
    
    If Not Vazio(Trim(strDATA_JULGAMENTO_INI)) And Not Vazio(Trim(strDATA_JULGAMENTO_FIM)) Then
    
        If DateValue(strDATA_JULGAMENTO_INI) > DateValue(strDATA_JULGAMENTO_FIM) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_RI_Registro_Insatisfacao"), "Data Fim Julgamento deve conter valor maior ou igual Data In cio Julgamento."
        End If
        
    End If

    strConnect = ConnectionSQL("IFV")
        
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   RI.NUM_RI  , " & Chr(10)
    Sql = Sql & GetData("RI.DATA_INCLUSAO") & " DATA_INCLUSAO , " & Chr(10)
    Sql = Sql & GetData("RI.DATA_JULGAMENTO") & " DATA_JULGAMENTO , " & Chr(10)
    Sql = Sql & GetData("RI.DATA_RESPOSTA_CLI") & " DATA_RESPOSTA_CLI , " & Chr(10)
    Sql = Sql & "   RI.CLASSIFICACAO_RI  , " & Chr(10)
    Sql = Sql & "   RI.CAUSA_INSATISF  , " & Chr(10)
    Sql = Sql & "   RI.COD_SETOR  , " & Chr(10)
    Sql = Sql & "   RI.DEVOLUCAO_PROD, " & Chr(10)
    Sql = Sql & "   RI.JULGMENTO, " & Chr(10)
    Sql = Sql & "   RI.STATUS_RI, " & Chr(10)
    Sql = Sql & "   RI.CLIENTE,   " & Chr(10)
    Sql = Sql & "   RI_SET.DES_SETOR,   " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)
    
    Sql = Sql & "   FROM RI_REGISTRO_INSATISFACAO RI " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = RI.CLIENTE) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN RI_SETORES RI_SET " & Chr(10)
    Sql = Sql & "        ON (RI_SET.COD_SETOR = RI.COD_SETOR) " & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strNUM_RI)) Then
         Sql = Sql & SqlAux & " RI.NUM_RI = '" & Substitui(strNUM_RI) & "'" & Chr(10)
         SqlAux = " AND "
    End If
            
    If Not Vazio(Trim(strLIKE_NUM_RI)) Then
         Sql = Sql & SqlAux & " RI.NUM_RI LIKE '%" & Substitui(strLIKE_NUM_RI) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_INCLUSAO)) Then
         Sql = Sql & SqlAux & " RI.DATA_INCLUSAO = " & ToData(strDATA_INCLUSAO) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_JULGAMENTO)) Then
         Sql = Sql & SqlAux & " RI.DATA_JULGAMENTO = " & ToData(strDATA_JULGAMENTO) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_INCLUSAO_INI)) Then
         Sql = Sql & SqlAux & " RI.DATA_INCLUSAO >= " & ToData(strDATA_INCLUSAO_INI) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_INCLUSAO_FIM)) Then
         Sql = Sql & SqlAux & " RI.DATA_INCLUSAO <= " & ToData(strDATA_INCLUSAO_FIM) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_JULGAMENTO_INI)) Then
         Sql = Sql & SqlAux & " RI.DATA_JULGAMENTO >= " & ToData(strDATA_JULGAMENTO_INI) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_JULGAMENTO_FIM)) Then
         Sql = Sql & SqlAux & " RI.DATA_JULGAMENTO <= " & ToData(strDATA_JULGAMENTO_FIM) & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strCOD_SETOR)) Then
         Sql = Sql & SqlAux & " RI.COD_SETOR = '" & Substitui(strCOD_SETOR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " RI.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
         
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If

    
    If Not Vazio(Trim(strSTATUS_RI)) Then
         Sql = Sql & SqlAux & " RI.STATUS_RI LIKE '%" & Substitui(strSTATUS_RI) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY RI.NUM_RI " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_RI_Registro_Insatisfacao = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_RI_Registro_Insatisfacao"), Err.Description

End Function




'
' Descric o : Listar_RI_Detalhes
' Retorno   : RecordSet
'

Public Function Listar_RI_Detalhes(Optional ByVal strNUM_RI As String, _
                                   Optional ByVal strLIKE_NUM_RI As String, _
                                   Optional ByVal strCLIENTE As String, _
                                   Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
        
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   RI_DET.NUM_RI,  " & Chr(10)
    Sql = Sql & "   RI_DET.ID_SEQ , " & Chr(10)
    Sql = Sql & GetData("RI_DET.DATA_OBS") & " DATA_OBS , " & Chr(10)
    Sql = Sql & "   RI_DET.USUARIO_OBS, " & Chr(10)
    Sql = Sql & "   RI_DET.DESCRICAO_OBS , " & Chr(10)
    Sql = Sql & "   RI.CLIENTE , " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)
    
    Sql = Sql & "   FROM RI_DETALHES RI_DET" & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN RI_REGISTRO_INSATISFACAO RI" & Chr(10)
    Sql = Sql & "     ON(RI.NUM_RI = RI_DET.NUM_RI)" & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = RI.CLIENTE) " & Chr(10)
   
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strNUM_RI)) Then
         Sql = Sql & SqlAux & " RI_DET.NUM_RI = '" & Substitui(strNUM_RI) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strLIKE_NUM_RI)) Then
         Sql = Sql & SqlAux & " RI.NUM_RI LIKE '%" & Substitui(strLIKE_NUM_RI) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " RI.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
         
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
        
    Sql = Sql & "ORDER BY RI_DET.NUM_RI,RI_DET.ID_SEQ " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_RI_Detalhes = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_RI_Detalhes"), Err.Description

End Function


'
' Descric o : Listar_RI_Setor
' Retorno   : RecordSet
'

Public Function Listar_RI_Setor(Optional ByVal strCOD_SETOR As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
        
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   RI_SET.COD_SETOR , " & Chr(10)
    Sql = Sql & "   RI_SET.DES_SETOR  " & Chr(10)
    
    Sql = Sql & "   FROM RI_SETORES RI_SET" & Chr(10)
   
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_SETOR)) Then
         Sql = Sql & SqlAux & " RI_SET.COD_SETOR = '" & Substitui(strCOD_SETOR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
  
        
    Sql = Sql & "ORDER BY RI_SET.DES_SETOR " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_RI_Setor = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_RI_Setor"), Err.Description

End Function

'
' Descric o : Listar_Membros
' Retorno   : RecordSet
'

Public Function Listar_Membros(Optional ByVal strCOD_MEMBRO As String, _
                               Optional ByVal strSENHA As String, _
                               Optional ByVal strSITUACAO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   COD_MEMBRO," & Chr(10)
    Sql = Sql & "   NOME_MEMBRO," & Chr(10)
    Sql = Sql & "   SENHA," & Chr(10)
    Sql = Sql & "   TIPO_MEMBRO," & Chr(10)
    Sql = Sql & "   SITUACAO," & Chr(10)
    Sql = Sql & "   EMAIL" & Chr(10)

    Sql = Sql & "   FROM MEMBROS " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_MEMBRO)) Then
         Sql = Sql & SqlAux & " COD_MEMBRO = '" & Substitui(strCOD_MEMBRO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSENHA)) Then
         Sql = Sql & SqlAux & " SENHA = '" & Substitui(strSENHA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSITUACAO)) Then
         Sql = Sql & SqlAux & " SITUACAO = '" & Substitui(strSITUACAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY COD_MEMBRO" & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Membros = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Membros"), Err.Description

End Function


## IFV01_FABRICA.bas
Attribute VB_Name = "IFV01_FABRICA"
Private Const strNomeModulo = "IFV01_FABRICA"

'
' Descric o : Listar_Fabrica
' Retorno   : RecordSet
'

Public Function Listar_Fabrica(Optional ByVal strCOD_FABRICA As String, _
                               Optional ByVal strLIKE_COD_FABRICA_DES_FABRICA As String) As ADODB.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As ADODB.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
 
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   FABRICA.COD_FABRICA , " & Chr(10)
    Sql = Sql & "   FABRICA.DES_FABRICA  " & Chr(10)
    
    Sql = Sql & "   FROM FABRICA FABRICA " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CLIENTE.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_FABRICA)) Then
         Sql = Sql & SqlAux & " FABRICA.COD_FABRICA = '" & Substitui(strCOD_FABRICA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_FABRICA_DES_FABRICA)) Then
         Sql = Sql & SqlAux & " (FABRICA.COD_FABRICA LIKE '%" & Substitui(strLIKE_COD_FABRICA_DES_FABRICA) & "%'" & Chr(10)
         Sql = Sql & " OR FABRICA.DES_FABRICA LIKE '%" & Substitui(strLIKE_COD_FABRICA_DES_FABRICA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY FABRICA.DES_FABRICA " & Chr(10)
           
    Set Rst = New ADODB.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Fabrica = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Fabrica"), Err.Description

End Function

'
' Descric o : Listar_Estoque_Produto
' Retorno   : RecordSet
'

Public Function Listar_Estoque_Produto(Optional ByVal strCOD_FABRICA As String, _
                                       Optional ByVal strLIKE_COD_FABRICA_DES_FABRICA As String, _
                                       Optional ByVal strCOD_PRODUTO As String, _
                                       Optional ByVal strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA As String, _
                                       Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                       Optional ByVal strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO As String, _
                                       Optional ByVal strCOD_QUALIDADE As String, _
                                       Optional ByVal strLIKE_COD_QUALIDADE_DES_QUALIDADE As String, _
                                       Optional ByVal strVISAO As String, _
                                       Optional ByVal strTIPO_DE_ESTOQUE As String) As ADODB.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As ADODB.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    If strVISAO = "E" Then
        Sql = Sql & "   EST_PROD.COD_FABRICA , " & Chr(10)
        Sql = Sql & "   FABRICA.DES_FABRICA,  " & Chr(10)
    Else
        Sql = Sql & "   PROD.COD_DIMENSAO , " & Chr(10)
        Sql = Sql & "   PROD.COD_COR , " & Chr(10)
        Sql = Sql & "   PROD.COD_EMBALAGEM , " & Chr(10)
    End If
    Sql = Sql & "   EST_PROD.COD_PRODUTO , " & Chr(10)
    
    
    Sql = Sql & "   ISNULL(SUM(EST_PROD.QTD_DISPONIVEL),'0') QTD_DISPONIVEL , " & Chr(10)
    Sql = Sql & "   ISNULL(SUM(EST_PROD.QTD_PLANEJ_POS),'0') QTD_PLANEJ_POS , " & Chr(10)
    Sql = Sql & "   ISNULL(SUM(EST_PROD.QTD_PLANEJ_NEG),'0')  QTD_PLANEJ_NEG, " & Chr(10)
    Sql = Sql & "   ISNULL(SUM(EST_PROD.QTD_VENDIDA),'0')  QTD_VENDIDA, " & Chr(10)
    Sql = Sql & "   (ISNULL(SUM(EST_PROD.QTD_DISPONIVEL),'0')+ISNULL(SUM(EST_PROD.QTD_PLANEJ_POS),'0')-ISNULL(SUM(EST_PROD.QTD_PLANEJ_NEG),'0'))  SALDO_P , " & Chr(10)
    Sql = Sql & "   (ISNULL(SUM(EST_PROD.QTD_DISPONIVEL),'0')-ISNULL(SUM(EST_PROD.QTD_VENDIDA),'0')) SALDO  , " & Chr(10)
    
    Sql = Sql & "   EST_PROD.TIPO_DE_ESTOQUE, " & Chr(10)
    
    Sql = Sql & "   PROD.COD_QUALIDADE , " & Chr(10)
    Sql = Sql & "   PROD.COD_ARTIGO_PADRAO  " & Chr(10)
    
    Sql = Sql & "   FROM ESTOQUE_PRODUTO EST_PROD " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN FABRICA FABRICA " & Chr(10)
    Sql = Sql & "        ON (FABRICA.COD_FABRICA = EST_PROD.COD_FABRICA) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN PRODUTO PROD " & Chr(10)
    Sql = Sql & "        ON (PROD.COD_PRODUTO = EST_PROD.COD_PRODUTO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN ARTIGO_PADRAO ART_PAD " & Chr(10)
    Sql = Sql & "        ON (PROD.COD_ARTIGO_PADRAO = ART_PAD.COD_ARTIGO_PADRAO) " & Chr(10)
    
    
   
    SqlAux = " WHERE " & Chr(10)
    
    
    If Not Vazio(Trim(strCOD_FABRICA)) Then
         Sql = Sql & SqlAux & " EST_PROD.COD_FABRICA = '" & Substitui(strCOD_FABRICA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_FABRICA_DES_FABRICA)) Then
         Sql = Sql & SqlAux & "( EST_PROD.COD_FABRICA LIKE '%" & Substitui(strLIKE_COD_FABRICA_DES_FABRICA) & "%'" & Chr(10)
         Sql = Sql & " OR  FABRICA.DES_FABRICA LIKE '%" & Substitui(strLIKE_COD_FABRICA_DES_FABRICA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strCOD_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.COD_PRODUTO = '" & Substitui(strCOD_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA)) Then
         Sql = Sql & SqlAux & "( PROD.COD_PRODUTO LIKE '%" & Substitui(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_PRODUTO_CURTA LIKE '%" & Substitui(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    
     If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " PROD.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & "( PROD.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%'" & Chr(10)
         Sql = Sql & " OR  ART_PAD.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strTIPO_DE_ESTOQUE)) Then
         Sql = Sql & SqlAux & " EST_PROD.TIPO_DE_ESTOQUE = '" & Substitui(strTIPO_DE_ESTOQUE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strCOD_QUALIDADE)) Then
         Sql = Sql & SqlAux & " PROD.COD_QUALIDADE = '" & Substitui(strCOD_QUALIDADE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_QUALIDADE_DES_QUALIDADE)) Then
         Sql = Sql & SqlAux & "( PROD.COD_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
            
    If strVISAO = "E" Then
                
        Sql = Sql & "  GROUP BY EST_PROD.COD_FABRICA,  " & Chr(10)
        Sql = Sql & "           FABRICA.DES_FABRICA,  " & Chr(10)
        Sql = Sql & "           EST_PROD.COD_PRODUTO,  " & Chr(10)
        Sql = Sql & "           EST_PROD.TIPO_DE_ESTOQUE, " & Chr(10)
        Sql = Sql & "           PROD.COD_QUALIDADE ,  " & Chr(10)
        Sql = Sql & "           PROD.COD_ARTIGO_PADRAO   " & Chr(10)
        
              
    
    Else
            
        Sql = Sql & "  GROUP BY EST_PROD.COD_FABRICA,  " & Chr(10)
        Sql = Sql & "           FABRICA.DES_FABRICA,  " & Chr(10)
        Sql = Sql & "           EST_PROD.COD_PRODUTO,  " & Chr(10)
        Sql = Sql & "           EST_PROD.TIPO_DE_ESTOQUE, " & Chr(10)
        Sql = Sql & "           PROD.COD_QUALIDADE ,  " & Chr(10)
        Sql = Sql & "           PROD.COD_ARTIGO_PADRAO,   " & Chr(10)
        Sql = Sql & "           PROD.COD_ARTIGO_PADRAO , " & Chr(10)
        Sql = Sql & "           PROD.COD_DIMENSAO , " & Chr(10)
        Sql = Sql & "           PROD.COD_COR , " & Chr(10)
        Sql = Sql & "           PROD.DES_QUALIDADE , " & Chr(10)
        Sql = Sql & "           PROD.COD_EMBALAGEM , " & Chr(10)
        Sql = Sql & "           PROD.COD_QUALIDADE  " & Chr(10)
    End If
    
    'sql = sql & "ORDER BY PROD.COD_PRODUTO " & Chr(10)
           
    Set Rst = New ADODB.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Estoque_Produto = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Estoque_Produto"), Err.Description

End Function



## IFV01_PEDIDO.bas
Attribute VB_Name = "IFV01_PEDIDO"
Private Const strNomeModulo = "IFV01_PEDIDO"

'
' Descric o : Listar_Pedido
' Retorno   : RecordSet
'

Public Function Listar_Pedido(Optional ByVal strNUM_PEDIDO As String, _
                              Optional ByVal strCOD_NOTA_FISCAL As String, _
                              Optional ByVal strCOD_SIT_PEDIDO As String, _
                              Optional ByVal strCLIENTE As String, _
                              Optional ByVal strDATA_INICIAL As String, _
                              Optional ByVal strDATA_FINAL As String, _
                              Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                              Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String


    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Pedido"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT DISTINCT" & Chr(10)
    Sql = Sql & "   PED.NUM_PEDIDO, " & Chr(10)
    Sql = Sql & "   PED.ESTABELECIMENTO, " & Chr(10)
    Sql = Sql & "   PED.CLIENTE, " & Chr(10)
    Sql = Sql & GetData("PED.DATA_INCLUSAO") & " DATA_INCLUSAO , " & Chr(10)
    Sql = Sql & "   PED.DES_TRANSPORTE, " & Chr(10)
    Sql = Sql & GetData("PED.DATA_ULT_ALT") & " DATA_ULT_ALT  , " & Chr(10)
    Sql = Sql & "   PED.PEDIDO_CLIENTE, " & Chr(10)
    Sql = Sql & "   PED.POSTO, " & Chr(10)
    Sql = Sql & "   ISNULL(PED.TOTAL_PEDIDO, 0 ) TOTAL_PEDIDO, " & Chr(10)
    Sql = Sql & "   PED.DES_COND_COMERCIAL, " & Chr(10)
    Sql = Sql & "   PED.DES_SIT_FATUR, " & Chr(10)
    Sql = Sql & "   PED.COD_SIT_PEDIDO, " & Chr(10)
    Sql = Sql & "   PED.DES_SIT_PEDIDO, " & Chr(10)
    Sql = Sql & GetData("PED.DATA_REQUERIDA") & " DATA_REQUERIDA  , " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)
        
    Sql = Sql & "   FROM PEDIDO PED " & Chr(10)
    
    Sql = Sql & "   INNER JOIN ITENS_PEDIDO ITENS_PED " & Chr(10)
    Sql = Sql & "         ON (ITENS_PED.NUM_PEDIDO = PED.NUM_PEDIDO) " & Chr(10)
    
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = PED.CLIENTE) " & Chr(10)
   
    If Not Vazio(Trim(strUN_Negocio)) Then
   
         Sql = Sql & " INNER JOIN ATIVIDADES_CLIENTE ATIV_CLIENTE" & Chr(10)
         Sql = Sql & "       ON(CLIENTE.CLIENTE = ATIV_CLIENTE.CLIENTE " & Chr(10)
         Sql = Sql & "       AND ATIV_CLIENTE.UNIDADE_NEGOCIO = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
         
    End If

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " PED.NUM_PEDIDO LIKE '%" & Substitui(strNUM_PEDIDO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_NOTA_FISCAL)) Then
         Sql = Sql & SqlAux & " PED.NUM_PEDIDO IN(SELECT NUM_PEDIDO FROM NOTA_FISCAL " & Chr(10)
         Sql = Sql & " WHERE COD_NOTA_FISCAL = '" & Substitui(strCOD_NOTA_FISCAL) & "')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_SIT_PEDIDO)) Then
         Sql = Sql & SqlAux & " PED.COD_SIT_PEDIDO = '" & Substitui(strCOD_SIT_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " PED.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_INICIAL)) Then
         Sql = Sql & SqlAux & " ITENS_PED.DATA_CEDO >= " & ToData(strDATA_INICIAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_FINAL)) Then
         Sql = Sql & SqlAux & " ITENS_PED.DATA_TARDE <= " & ToData(strDATA_FINAL) & Chr(10)
         SqlAux = " AND "
    End If
    
'---------- SE DATA FOR RETORNAR PARA PESQUISA POR PEDIDO
 '   If Not Vazio(Trim(strDATA_INICIAL)) Then
 '        Sql = Sql & SqlAux & " PED.DATA_INCLUSAO >= " & ToData(strDATA_INICIAL) & Chr(10)
 '        SqlAux = " AND "
 '   End If
 '
 '   If Not Vazio(Trim(strDATA_FINAL)) Then
 '        Sql = Sql & SqlAux & " PED.DATA_INCLUSAO <= " & ToData(strDATA_FINAL) & Chr(10)
 '        SqlAux = " AND "
 '   End If
 '-----------------------------------
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY PED.NUM_PEDIDO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Pedido = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Pedido"), Err.Description

End Function
'---------------------------------------------------------------


'
' Descric o : Listar_Nota_Fi
' Retorno   : RecordSet
'

Public Function Listar_ND(Optional ByVal strCNPJ) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    If Vazio(Trim(strCNPJ)) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_NF"), "CNPJ em Branco"
        
    End If

    strConnect = ConnectionSQL("IFV")


    Sql = "SELECT"
    Sql = Sql & " BUSINESS_UNIT "
    Sql = Sql & ",CLIENTE "
    Sql = Sql & ",TITULO "
    Sql = Sql & ",SEQ "
    Sql = Sql & ",TIPO "
    Sql = Sql & ",MOTIVO "
    Sql = Sql & ",DESCRICAO "
    Sql = Sql & ",EMISSAO "
    Sql = Sql & ",VCTO "
    Sql = Sql & ",SALDO "
    Sql = Sql & ",VL_ORIGINAL "
    Sql = Sql & "From dbo.ND "
    Sql = Sql & " WHERE CLIENTE = '" & Mid(RTrim(strCNPJ), 1, 8) & "' "
    
    'SqlAux = " WHERE " & Chr(10)
    

    'If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
    '     Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
    '     Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
    '     SqlAux = " AND "
    'End If
        
    Sql = Sql & "ORDER BY CLIENTE, TITULO  " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_ND = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_ND"), Err.Description

End Function


'---------------------------------------------------------------


'
' Descric o : Listar_Nota_Fiscal
' Retorno   : RecordSet
'

Public Function Listar_Nota_Fiscal(Optional ByVal strCOD_NOTA_FISCAL As String, _
                                   Optional ByVal strSERIE As String, _
                                   Optional ByVal strNUM_PEDIDO As String, _
                                   Optional ByVal strCLIENTE As String, _
                                   Optional ByVal strDATA_INICIAL As String, _
                                   Optional ByVal strDATA_FINAL As String, _
                                   Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Pedido"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   NOT_FIS.COD_NOTA_FISCAL , " & Chr(10)
    Sql = Sql & "   NOT_FIS.SERIE , " & Chr(10)
    Sql = Sql & "   NOT_FIS.NUM_PEDIDO , " & Chr(10)
    Sql = Sql & "   NOT_FIS.CLIENTE , " & Chr(10)
    Sql = Sql & "   NOT_FIS.ESTABELECIMENTO , " & Chr(10)
    Sql = Sql & "   NOT_FIS.COD_FABRICA , " & Chr(10)
    Sql = Sql & "   NOT_FIS.STATUS_NF , " & Chr(10)
    Sql = Sql & "   NOT_FIS.TIPO_NF , " & Chr(10)
    Sql = Sql & GetData("NOT_FIS.DATA_EMISSAO") & " DATA_EMISSAO , " & Chr(10)
    Sql = Sql & GetData("NOT_FIS.DATA_SAIDA_MER") & " DATA_SAIDA_MER , " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_BCICM,'0') VALOR_BCICM, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_ICM,'0') VALOR_ICM, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_IPI,'0') VALOR_IPI, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_ALIQICM,'0') VALOR_ALIQICM, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.PESO_LIQ,'0') PESO_LIQ, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.PESO_BRUTO,'0') PESO_BRUTO, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_DESC,'0') VALOR_DESC, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_TOTAL,'0') VALOR_TOTAL, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.TOTAL_UNID_FATUR,'0') TOTAL_UNID_FATUR, " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.QTD_VOLUME,'0') QTD_VOLUME, " & Chr(10)
    Sql = Sql & "   NOT_FIS.VIA_TRANSPORTE , " & Chr(10)
    Sql = Sql & "   NOT_FIS.DES_TRANSPORTE , " & Chr(10)
    Sql = Sql & "   ISNULL(NOT_FIS.VALOR_DESC_PONT,'0') VALOR_DESC_PONT, " & Chr(10)
    Sql = Sql & "   NOT_FIS.DES_QUALIDADE , " & Chr(10)
    Sql = Sql & "   NOT_FIS.CODMOEDA , " & Chr(10)
    Sql = Sql & "   FABR.DES_FABRICA, " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA, " & Chr(10)
    Sql = Sql & "   NOT_FIS.COD_TRANSP " & Chr(10)

    Sql = Sql & "   FROM NOTA_FISCAL NOT_FIS " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN FABRICA FABR " & Chr(10)
    Sql = Sql & "        ON (FABR.COD_FABRICA = NOT_FIS.COD_FABRICA) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = NOT_FIS.CLIENTE) " & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_NOTA_FISCAL)) Then
         Sql = Sql & SqlAux & " NOT_FIS.COD_NOTA_FISCAL LIKE '%" & Substitui(strCOD_NOTA_FISCAL) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSERIE)) Then
         Sql = Sql & SqlAux & " NOT_FIS.SERIE LIKE '%" & Substitui(strSERIE) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " NOT_FIS.NUM_PEDIDO LIKE '%" & Substitui(strNUM_PEDIDO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
     If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " NOT_FIS.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_INICIAL)) Then
         Sql = Sql & SqlAux & " NOT_FIS.DATA_EMISSAO >= " & ToData(strDATA_INICIAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_FINAL)) Then
         Sql = Sql & SqlAux & " NOT_FIS.DATA_EMISSAO <= " & ToData(strDATA_FINAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY NOT_FIS.COD_FABRICA, NOT_FIS.COD_NOTA_FISCAL, NOT_FIS.SERIE  " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Nota_Fiscal = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Nota_Fiscal"), Err.Description

End Function

'
' Descric o : Listar_Itens_Nota_Fiscal
' Retorno   : RecordSet
'

Public Function Listar_Itens_Nota_Fiscal(Optional ByVal strCOD_NOTA_FISCAL As String, _
                                         Optional ByVal strSERIE As String, _
                                         Optional ByVal strID_SEQUENCIAL As String, _
                                         Optional ByVal strNUM_PEDIDO As String, _
                                         Optional ByVal strCOD_FABRICA As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   ITENS_NOT_FIS.COD_NOTA_FISCAL , " & Chr(10)
    Sql = Sql & "   ITENS_NOT_FIS.SERIE , " & Chr(10)
    Sql = Sql & "   ITENS_NOT_FIS.ID_SEQUENCIAL , " & Chr(10)
    Sql = Sql & "   ITENS_NOT_FIS.COD_PRODUTO , " & Chr(10)
    Sql = Sql & "   ISNULL(ITENS_NOT_FIS.QTD_COMER,'0') QTD_COMER, " & Chr(10)
    Sql = Sql & "   ISNULL(ITENS_NOT_FIS.QTD_FATURADA,'0') QTD_FATURADA, " & Chr(10)
    Sql = Sql & "   ISNULL(ITENS_NOT_FIS.ALIQ_IPI,'0') ALIQ_IPI, " & Chr(10)
    Sql = Sql & "   ISNULL(ITENS_NOT_FIS.PREC_UNITARIO,'0') PREC_UNITARIO, " & Chr(10)
    Sql = Sql & "   ISNULL(ITENS_NOT_FIS.PESO_LIQ,'0') PESO_LIQ, " & Chr(10)
    Sql = Sql & "   ITENS_NOT_FIS.POSFISC , " & Chr(10)
    Sql = Sql & "   NOT_FIS.NUM_PEDIDO, " & Chr(10)
    Sql = Sql & "   POD.DES_PRODUTO_CURTA, " & Chr(10)
    Sql = Sql & "   POD.DES_PRODUTO_LONGA " & Chr(10)

    Sql = Sql & "   FROM ITENS_NOTA_FISCAL ITENS_NOT_FIS " & Chr(10)
    Sql = Sql & "   INNER JOIN NOTA_FISCAL NOT_FIS" & Chr(10)
    Sql = Sql & "       ON(NOT_FIS.COD_FABRICA = ITENS_NOT_FIS.COD_FABRICA " & Chr(10)
    Sql = Sql & "          AND NOT_FIS.COD_NOTA_FISCAL = ITENS_NOT_FIS.COD_NOTA_FISCAL " & Chr(10)
    Sql = Sql & "          AND NOT_FIS.SERIE = ITENS_NOT_FIS.SERIE )" & Chr(10)
    Sql = Sql & "   INNER JOIN PRODUTO POD" & Chr(10)
    Sql = Sql & "       ON(POD.COD_PRODUTO = ITENS_NOT_FIS.COD_PRODUTO ) " & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_FABRICA)) Then
         Sql = Sql & SqlAux & " ITENS_NOT_FIS.COD_FABRICA = '" & Substitui(strCOD_FABRICA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_NOTA_FISCAL)) Then
         Sql = Sql & SqlAux & " ITENS_NOT_FIS.COD_NOTA_FISCAL = '" & Substitui(strCOD_NOTA_FISCAL) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSERIE)) Then
         Sql = Sql & SqlAux & " ITENS_NOT_FIS.SERIE = '" & Substitui(strSERIE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strID_SEQUENCIAL)) Then
         Sql = Sql & SqlAux & " ITENS_NOT_FIS.ID_SEQUENCIAL = " & strID_SEQUENCIAL & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " NOT_FIS.NUM_PEDIDO = '" & Substitui(strNUM_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
            
    Sql = Sql & "ORDER BY NOT_FIS.COD_NOTA_FISCAL, ITENS_NOT_FIS.SERIE, ITENS_NOT_FIS.ID_SEQUENCIAL  " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Itens_Nota_Fiscal = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Itens_Nota_Fiscal"), Err.Description

End Function


'
' Descric o : Listar_Duplicatas
' Retorno   : RecordSet
'

Public Function Listar_Duplicatas(Optional ByVal strCOD_NOTA_FISCAL As String, _
                                  Optional ByVal strSERIE As String, _
                                  Optional ByVal strSEQUENCIA As String, _
                                  Optional ByVal strNUM_PEDIDO As String, _
                                  Optional ByVal strCLIENTE As String, _
                                  Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                                  Optional ByVal strDATA_VENCIMENTO_INICIAL As String, _
                                  Optional ByVal strDATA_VENCIMENTO_FINAL As String, _
                                  Optional ByVal strDATA_PAGAMENTO_INICIAL As String, _
                                  Optional ByVal strDATA_PAGAMENTO_FINAL As String, _
                                  Optional ByVal strStatus As String, _
                                  Optional ByVal strTpLancto As Integer, _
                                  Optional ByVal strUN_Negocio As String, _
                                  Optional ByVal strCOD_FABRICA As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    
    If Not Vazio(Trim(strDATA_VENCIMENTO_INICIAL)) And Not Vazio(Trim(strDATA_VENCIMENTO_FINAL)) Then
    
        If DateValue(strDATA_VENCIMENTO_INICIAL) > DateValue(strDATA_VENCIMENTO_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Pedido"), "Data Vencimento Final deve conter valor maior ou igual Data Vencimento In cio."
        End If
        
    End If
    
    If Not Vazio(Trim(strDATA_PAGAMENTO_INICIAL)) And Not Vazio(Trim(strDATA_PAGAMENTO_FINAL)) Then
    
        If DateValue(strDATA_PAGAMENTO_INICIAL) > DateValue(strDATA_PAGAMENTO_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Pedido"), "Data Pagamento Final deve conter valor maior ou igual Data Pagamento In cio."
        End If
        
    End If

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT DISTINCT " & Chr(10)
    Sql = Sql & "   DUPL.COD_NOTA_FISCAL , " & Chr(10)
    Sql = Sql & "   DUPL.SERIE , " & Chr(10)
    Sql = Sql & "   DUPL.SEQUENCIA, " & Chr(10)
    Sql = Sql & GetData("DUPL.DATA_VENCIMENTO") & " DATA_VENCIMENTO , " & Chr(10)
    Sql = Sql & "   DUPL.NUM_TITULO , " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR,'0') VALOR, " & Chr(10)
    Sql = Sql & "   DUPL.DES_PORTADOR , " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR_PAGAMENTO,'0') VALOR_PAGAMENTO, " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR_OUTROS,'0') VALOR_OUTROS, " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR_SALDO,'0') VALOR_SALDO, " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR_COMISSAO,'0') VALOR_COMISSAO, " & Chr(10)
    Sql = Sql & "   DUPL.TIPO_LCTO , " & Chr(10)
    Sql = Sql & GetData("DUPL.DATA_PAGAMENTO") & " DATA_PAGAMENTO , " & Chr(10)
    Sql = Sql & "   NOT_FIS.NUM_PEDIDO, " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)

    Sql = Sql & "   FROM DUPLICATAS DUPL " & Chr(10)
    Sql = Sql & "   INNER JOIN NOTA_FISCAL NOT_FIS" & Chr(10)
    Sql = Sql & "       ON(NOT_FIS.COD_FABRICA  = DUPL.COD_FABRICA " & Chr(10)
    Sql = Sql & "          AND NOT_FIS.COD_NOTA_FISCAL = DUPL.COD_NOTA_FISCAL " & Chr(10)
    Sql = Sql & "          AND NOT_FIS.SERIE = DUPL.SERIE )" & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = NOT_FIS.CLIENTE) " & Chr(10)
    
    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & " INNER JOIN ATIVIDADES_CLIENTE ATIV_CLIENTE" & Chr(10)
         Sql = Sql & " ON(CLIENTE.CLIENTE = ATIV_CLIENTE.CLIENTE " & Chr(10)
         Sql = Sql & " AND ATIV_CLIENTE.UNIDADE_NEGOCIO = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
    End If
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_FABRICA)) Then
         Sql = Sql & SqlAux & " DUPL.COD_FABRICA = '" & Substitui(strCOD_FABRICA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_NOTA_FISCAL)) Then
         Sql = Sql & SqlAux & " DUPL.COD_NOTA_FISCAL = '" & Substitui(strCOD_NOTA_FISCAL) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSERIE)) Then
         Sql = Sql & SqlAux & " DUPL.SERIE = '" & Substitui(strSERIE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSEQUENCIA)) Then
         Sql = Sql & SqlAux & " DUPL.SEQUENCIA = '" & Substitui(strSEQUENCIA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & SqlAux & " DUPL.UN_NEGOCIO_DUP = '" & Substitui(strUN_Negocio) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " NOT_FIS.NUM_PEDIDO = '" & Substitui(strNUM_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
     If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " NOT_FIS.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_VENCIMENTO_INICIAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_VENCIMENTO >= " & ToData(strDATA_VENCIMENTO_INICIAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_VENCIMENTO_FINAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_VENCIMENTO <= " & ToData(strDATA_VENCIMENTO_FINAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_PAGAMENTO_INICIAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO >= " & ToData(strDATA_PAGAMENTO_INICIAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_PAGAMENTO_FINAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO <= " & ToData(strDATA_PAGAMENTO_FINAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    Select Case strStatus
            '  VENCER
            Case "A"
                Sql = Sql & SqlAux & " DUPL.VALOR_SALDO <> 0 " & Chr(10)
                Sql = Sql & " AND DUPL.DATA_VENCIMENTO >= " & ToData(Date) & Chr(10)
                SqlAux = " AND "
            'VENCIDO
            Case "V"
                Sql = Sql & SqlAux & " DUPL.VALOR_SALDO <> 0 " & Chr(10)
                Sql = Sql & " AND DUPL.DATA_VENCIMENTO < " & ToData(Date) & Chr(10)
                SqlAux = " AND "
            'LIQUIDADO
            Case "L"
                Sql = Sql & SqlAux & " DUPL.VALOR_SALDO = 0 " & Chr(10)
                Sql = Sql & " AND DUPL.DATA_PAGAMENTO IS NOT NULL " & Chr(10)
                SqlAux = " AND "
    
    End Select
    strTpLancto = 7
    If strTpLancto > 0 And strTpLancto <= 6 Then
       Select Case strTpLancto
            'C/Devol / Aberto
            Case 6
                Sql = Sql & SqlAux & " DUPL.TIPO_LCTO IN ('C/Devol','Aberto') " & Chr(10)
                SqlAux = " AND "
            'Normal / Aberto
            Case 5
                Sql = Sql & SqlAux & " DUPL.TIPO_LCTO IN ('Normal','Aberto') " & Chr(10)
                SqlAux = " AND "
           'Aberto
           Case 4
                Sql = Sql & SqlAux & " DUPL.TIPO_LCTO = 'Aberto' " & Chr(10)
                SqlAux = " AND "
           'Normal / C/Devol
           Case 3
                Sql = Sql & SqlAux & " DUPL.TIPO_LCTO IN ('Normal','C/Devol') " & Chr(10)
                SqlAux = " AND "
           'C/Devol
           Case 2
                Sql = Sql & SqlAux & " DUPL.TIPO_LCTO = 'C/Devol' " & Chr(10)
                SqlAux = " AND "
           'Normal
           Case 1
                Sql = Sql & SqlAux & " DUPL.TIPO_LCTO = 'Normal' " & Chr(10)
                SqlAux = " AND "
       End Select
    End If
    
    Sql = Sql & "ORDER BY DUPL.COD_NOTA_FISCAL, DUPL.SERIE, DUPL.SEQUENCIA  " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Duplicatas = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Duplicatas"), Err.Description

End Function

'
' Descric o : Listar_Observacao_Pedido
' Retorno   : RecordSet
'

Public Function Listar_Observacao_Pedido(Optional ByVal strNUM_PEDIDO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   OBS_PED.NUM_PEDIDO, " & Chr(10)
    Sql = Sql & "   OBS_PED.ID_SEQUENCIAL, " & Chr(10)
    Sql = Sql & "   OBS_PED.DES_TIPO_OPERACAO, " & Chr(10)
    Sql = Sql & "   OBS_PED.INSCRICAO_ESTADUAL, " & Chr(10)
    Sql = Sql & "   OBS_PED.NOME, " & Chr(10)
    Sql = Sql & "   OBS_PED.ENDERECO, " & Chr(10)
    Sql = Sql & "   OBS_PED.MRH, " & Chr(10)
    Sql = Sql & "   OBS_PED.CIDADE, " & Chr(10)
    Sql = Sql & "   OBS_PED.UF, " & Chr(10)
    Sql = Sql & "   OBS_PED.MUNICIPIO, " & Chr(10)
    Sql = Sql & "   OBS_PED.TEXTO_NOTA_FISCAL, " & Chr(10)
    Sql = Sql & "   OBS_PED.TEXTO_LIVRE " & Chr(10)

    Sql = Sql & "   FROM OBSERVACAO_PEDIDO OBS_PED " & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)
        
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " OBS_PED.NUM_PEDIDO = '" & Substitui(strNUM_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY OBS_PED.NUM_PEDIDO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Observacao_Pedido = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Observacao_Pedido"), Err.Description

End Function

'
' Descric o : Listar_Situacao_Pedido
' Retorno   : RecordSet
'

Public Function Listar_Situacao_Pedido(Optional ByVal strCOD_SIT_PEDIDO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   DISTINCT PED.COD_SIT_PEDIDO, " & Chr(10)
    Sql = Sql & "   PED.DES_SIT_PEDIDO  " & Chr(10)
    Sql = Sql & "   FROM PEDIDO PED " & Chr(10)
    Sql = Sql & "   WHERE COD_SIT_PEDIDO IS NOT NULL " & Chr(10)
    Sql = Sql & "   AND DES_SIT_PEDIDO IS NOT NULL" & Chr(10)
    
    SqlAux = " AND " & Chr(10)
        
    If Not Vazio(Trim(strCOD_SIT_PEDIDO)) Then
         Sql = Sql & SqlAux & " PED.COD_SIT_PEDIDO = '" & Substitui(strCOD_SIT_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY PED.DES_SIT_PEDIDO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Situacao_Pedido = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Situacao_Pedido"), Err.Description

End Function

'
' Descric o : Listar_Itens_Pedido
' Retorno   : RecordSet
'

Public Function Listar_Itens_Pedido(Optional ByVal strNUM_PEDIDO As String, _
                                    Optional ByVal strCOD_PRODUTO As String, _
                                    Optional ByVal strID_SEQUENCIAL As String, _
                                    Optional ByVal strDATA_INICIAL As String, _
                                    Optional ByVal strDATA_FINAL As String, _
                                    Optional ByVal strQUALIDADE As String, _
                                    Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim SqlI           As String
    Dim RstIt          As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAuxI        As String

    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Itens_Pedido"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If
    
    strConnect = ConnectionSQL("IFV")

    SqlI = " SELECT " & Chr(10)
    
    SqlI = SqlI & "   ITENS_PED.NUM_PEDIDO , " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.COD_PRODUTO , " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.ID_SEQUENCIAL , " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.QTD_PEDIDA,'0') QTD_PEDIDA, " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.QTD_FATURADA,'0') QTD_FATURADA, " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.QTD_DESTINADA,'0') QTD_DESTINADA, " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.QTD_EMPENHADA,'0') QTD_EMPENHADA, " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.SITUACAO , " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.VLR_PRECO,'0') VLR_PRECO, " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.VLR_SALDO,'0') VLR_SALDO, " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.VLR_UNITARIO,'0') VLR_UNITARIO, " & Chr(10)
    SqlI = SqlI & "   ISNULL(ITENS_PED.DESCONTO,'0') DESCONTO, " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.COND_PAGTO, " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.TABELA_PRECO , " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.GRP_FISCAL_PRC, " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.OPER_FISCAL_PRC, " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.GRP_FISCAL_ENT, " & Chr(10)
    SqlI = SqlI & "   ITENS_PED.OPER_FISCAL_ENT, " & Chr(10)
    SqlI = SqlI & GetData("ITENS_PED.DATA_CEDO") & " DATA_CEDO  , " & Chr(10)
    SqlI = SqlI & GetData("ITENS_PED.DATA_TARDE") & " DATA_TARDE  , " & Chr(10)
    SqlI = SqlI & GetData("ITENS_PED.DATA_BASE_LN") & " DATA_BASE  , " & Chr(10)
    SqlI = SqlI & "   POD.DES_PRODUTO_CURTA, " & Chr(10)
    SqlI = SqlI & "   POD.DES_PRODUTO_LONGA, " & Chr(10)
    SqlI = SqlI & "   PED.COD_SIT_PEDIDO, " & Chr(10)
    SqlI = SqlI & "   CLIENTE.RAZAO_SOCIAL " & Chr(10)

    SqlI = SqlI & " FROM PEDIDO PED" & Chr(10)
    SqlI = SqlI & " INNER JOIN ITENS_PEDIDO ITENS_PED ON (ITENS_PED.NUM_PEDIDO = PED.NUM_PEDIDO)" & Chr(10)
    
    SqlI = SqlI & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    SqlI = SqlI & "        ON (CLIENTE.CLIENTE = PED.CLIENTE) " & Chr(10)
    
    SqlI = SqlI & " INNER JOIN PRODUTO POD" & Chr(10)
    
    If Not Vazio(Trim(strUN_Negocio)) Then
       SqlI = SqlI & "       ON(POD.COD_PRODUTO = ITENS_PED.COD_PRODUTO " & Chr(10)
       SqlI = SqlI & "       AND POD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
    Else
       SqlI = SqlI & "       ON(POD.COD_PRODUTO = ITENS_PED.COD_PRODUTO ) " & Chr(10)
    End If
    
    SqlAuxI = " WHERE "
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         SqlI = SqlI & SqlAuxI & " ITENS_PED.NUM_PEDIDO = '" & Substitui(strNUM_PEDIDO) & "'" & Chr(10)
         SqlAuxI = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_PRODUTO)) Then
       If Len(Trim(strCOD_PRODUTO)) = 17 Then
         SqlI = SqlI & SqlAuxI & " ITENS_PED.COD_PRODUTO = '" & Substitui(strCOD_PRODUTO) & "'" & Chr(10)
         SqlAuxI = " AND "
       Else
         SqlI = SqlI & SqlAuxI & " ITENS_PED.COD_PRODUTO LIKE '" & Substitui(strCOD_PRODUTO) & "%'" & Chr(10)
         SqlAuxI = " AND "
       End If
    End If
    
    If Not Vazio(Trim(strID_SEQUENCIAL)) Then
         SqlI = SqlI & SqlAuxI & " ITENS_PED.ID_SEQUENCIAL = " & strID_SEQUENCIAL & Chr(10)
         SqlAuxI = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
         SqlI = SqlI & SqlAuxI & "( ITENS_PED.DATA_CEDO >= " & ToData(strDATA_INICIAL) & Chr(10)
         SqlAuxI = " AND "
         SqlI = SqlI & SqlAuxI & " ITENS_PED.DATA_TARDE <= " & ToData(strDATA_FINAL) & " ) " & Chr(10)
         SqlAuxI = " AND "
    End If
    
    If Not Vazio(Trim(strQUALIDADE)) Then
         SqlI = SqlI & SqlAuxI & " POD.COD_QUALIDADE = '" & Substitui(strQUALIDADE) & "'" & Chr(10)
    End If

            
    SqlI = SqlI & "ORDER BY ITENS_PED.NUM_PEDIDO, ITENS_PED.COD_PRODUTO, ITENS_PED.ID_SEQUENCIAL  " & Chr(10)
           
    Set RstIt = New adodb.Recordset

    RstIt.CursorLocation = adUseClient
    RstIt.Open SqlI, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Itens_Pedido = RstIt

    Exit Function

ErrorHandler:

    If Not RstIt Is Nothing Then
        Set RstIt = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Itens_Pedido"), Err.Description

End Function



'
' Descric o : Listar_Pedido_Bloqueios
' Retorno   : RecordSet
'

Public Function Listar_Pedido_Bloqueios(Optional ByVal strNUM_PEDIDO As String, _
                                        Optional ByVal strNUM_LINHA As String, _
                                        Optional ByVal strID_SEQUENCIAL As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PED_BLOQ.NUM_PEDIDO , " & Chr(10)
    Sql = Sql & "   PED_BLOQ.NUM_LINHA, " & Chr(10)
    Sql = Sql & "   PED_BLOQ.ID_SEQUENCIAL , " & Chr(10)
    Sql = Sql & "   PED_BLOQ.DES_BLOQUEIO , " & Chr(10)
    Sql = Sql & "   PED_BLOQ.STATUS , " & Chr(10)
    Sql = Sql & "   PED_BLOQ.DATA_STATUS , " & Chr(10)
    Sql = Sql & "   PED_BLOQ.DES_MENSAGEM,  " & Chr(10)
    Sql = Sql & "   PED_BLOQ.ID_BLOQUEIO  " & Chr(10)
    
    Sql = Sql & "   FROM PEDIDO_BLOQUEIOS PED_BLOQ " & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " PED_BLOQ.NUM_PEDIDO = '" & Substitui(strNUM_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strNUM_LINHA)) Then
         Sql = Sql & SqlAux & " PED_BLOQ.NUM_LINHA = '" & Substitui(strNUM_LINHA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strID_SEQUENCIAL)) Then
         Sql = Sql & SqlAux & " PED_BLOQ.ID_SEQUENCIAL = " & strID_SEQUENCIAL & Chr(10)
         SqlAux = " AND "
    End If
            
    Sql = Sql & "ORDER BY PED_BLOQ.NUM_PEDIDO, PED_BLOQ.ID_SEQUENCIAL  " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Pedido_Bloqueios = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Pedido_Bloqueios"), Err.Description

End Function

' Descric o : Listar_ABC
' Retorno   : RecordSet
'

Public Function Listar_ABC(ByVal strANO As String, _
                           Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                           Optional ByVal strLINHA_PRODUTO As String, _
                           Optional ByVal strQUALIDADE As String, _
                           Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim RstAbc         As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String


    If Vazio(Trim(strANO)) Then
    
        Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_ABC"), "Ano deve ser Informado."
        
    End If

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   PROD.LINHA_PRODUTO, " & Chr(10)
    Sql = Sql & "   PROD.FAMILIA_COTA, " & Chr(10)
    Sql = Sql & "   PROD.DES_PRODUTO_CURTA, " & Chr(10)
    Sql = Sql & "   RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2) DATA_REQUERIDA, " & Chr(10)
    Sql = Sql & "   ISNULL(SUM(ITENS.QTD_PEDIDA), 0 ) TOTAL_PEDIDO " & Chr(10)
    
    Sql = Sql & "   FROM PEDIDO PED " & Chr(10)
    
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = PED.CLIENTE) " & Chr(10)
   
    Sql = Sql & "   LEFT OUTER JOIN ITENS_PEDIDO ITENS " & Chr(10)
    Sql = Sql & "        ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO) " & Chr(10)
   
    If Not Vazio(Trim(strUN_Negocio)) Then
   
         Sql = Sql & "   LEFT OUTER JOIN PRODUTO PROD " & Chr(10)
         Sql = Sql & "        ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO " & Chr(10)
         Sql = Sql & "        AND PROD.UN_NEGOCIO_PROD  = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
   
         Sql = Sql & " INNER JOIN ATIVIDADES_CLIENTE ATIV_CLIENTE" & Chr(10)
         Sql = Sql & "       ON(CLIENTE.CLIENTE = ATIV_CLIENTE.CLIENTE " & Chr(10)
         Sql = Sql & "       AND ATIV_CLIENTE.UNIDADE_NEGOCIO = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
    Else
    
       Sql = Sql & "   LEFT OUTER JOIN PRODUTO PROD " & Chr(10)
       Sql = Sql & "        ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO) " & Chr(10)

    End If
   
    SqlAux = " WHERE " & Chr(10)

    Sql = Sql & SqlAux & " PED.COD_SIT_PEDIDO <> 'C' " & Chr(10)
    
    SqlAux = " AND "
    
    If Not Vazio(Trim(strANO)) Then
         Sql = Sql & SqlAux & " YEAR(ITENS.DATA_CEDO) = '" & Substitui(strANO) & "'" & Chr(10)
    End If
    
    Sql = Sql & SqlAux & " ITENS.SITUACAO <> 'C' " & Chr(10)
    
    Sql = Sql & SqlAux & " ITENS.QTD_PEDIDA Is Not Null " & Chr(10)
    
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.LINHA_PRODUTO LIKE '" & Substitui(strLINHA_PRODUTO) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strQUALIDADE)) Then
         Sql = Sql & SqlAux & " PROD.COD_QUALIDADE = '" & Substitui(strQUALIDADE) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
    End If
        
    Sql = Sql & "GROUP BY CLIENTE.CNPJ, CLIENTE.RAZAO_SOCIAL, PROD.LINHA_PRODUTO, PROD.FAMILIA_COTA, PROD.DES_PRODUTO_CURTA, RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2) " & Chr(10)
    
    Sql = Sql & "ORDER BY CLIENTE.CNPJ, CLIENTE.RAZAO_SOCIAL, PROD.LINHA_PRODUTO, PROD.FAMILIA_COTA, PROD.DES_PRODUTO_CURTA, RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2) " & Chr(10)
    
    Set RstAbc = New adodb.Recordset

    RstAbc.CursorLocation = adUseClient
    RstAbc.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_ABC = RstAbc

    Exit Function

ErrorHandler:

    If Not RstAbc Is Nothing Then
        Set RstAbc = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_ABC"), Err.Description

End Function

' Descric o : Listar_EstFamCota
' Retorno   : RecordSet
'

Public Function Listar_EstFamCota(ByVal strDTINIC As String, _
                                  ByVal strDTFIM As String, _
                                  Optional ByVal strLINHA_PRODUTO As String, _
                                  Optional ByVal strQUALIDPROD As String, _
                                  Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim RstEstFam      As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String


    If Vazio(Trim(strDTINIC)) Then
    
        Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_EstFamCota"), "Mes Ano Inicial deve ser Informado."
        
    End If
    If Vazio(Trim(strDTFIM)) Then
    
        Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_EstFamCota"), "Mes Ano Final deve ser Informado."
        
    End If

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD.LINHA_PRODUTO, " & Chr(10)
    Sql = Sql & "   PROD.FAMILIA_COTA, " & Chr(10)
    Sql = Sql & "   RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2) MES_REQ, " & Chr(10)
    Sql = Sql & "   RIGHT('0' + LTRIM(YEAR(ITENS.DATA_CEDO)),4) ANO_REQ, " & Chr(10)
    Sql = Sql & "   ISNULL(SUM(ITENS.QTD_PEDIDA), 0 ) TOTAL_PEDIDO " & Chr(10)
    
    Sql = Sql & "   FROM PEDIDO PED " & Chr(10)
   
    Sql = Sql & "   LEFT OUTER JOIN ITENS_PEDIDO ITENS " & Chr(10)
    Sql = Sql & "        ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO) " & Chr(10)
   
    Sql = Sql & "   LEFT OUTER JOIN PRODUTO PROD " & Chr(10)
   
    If Not Vazio(Trim(strUN_Negocio)) Then
       Sql = Sql & "       ON(PROD.COD_PRODUTO = ITENS.COD_PRODUTO " & Chr(10)
       Sql = Sql & "       AND PROD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
    Else
       Sql = Sql & "       ON(PROD.COD_PRODUTO = ITENS.COD_PRODUTO) " & Chr(10)
    End If
   
    SqlAux = " WHERE " & Chr(10)

    Sql = Sql & SqlAux & " PED.COD_SIT_PEDIDO <> 'C' " & Chr(10)
    
    SqlAux = " AND "
    
    If Not Vazio(Trim(strDTINIC)) Then
       Sql = Sql & SqlAux & " ITENS.DATA_CEDO >= CONVERT(DATETIME, '" & Substitui(strDTINIC) & "',126)" & Chr(10)
    End If
    If Not Vazio(Trim(strDTFIM)) Then
       Sql = Sql & SqlAux & " ITENS.DATA_CEDO < CONVERT(DATETIME, '" & Substitui(strDTFIM) & "',126)" & Chr(10)
    End If
    
    Sql = Sql & SqlAux & " ITENS.SITUACAO <> 'C' " & Chr(10)
    
    Sql = Sql & SqlAux & " ITENS.QTD_PEDIDA Is Not Null " & Chr(10)
    
    If Not Vazio(Trim(strQUALIDPROD)) Then
         Sql = Sql & SqlAux & " PROD.COD_QUALIDADE = '" & Substitui(strQUALIDPROD) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.LINHA_PRODUTO LIKE '" & Substitui(strLINHA_PRODUTO) & "'" & Chr(10)
    End If
        
    Sql = Sql & "GROUP BY PROD.LINHA_PRODUTO, PROD.FAMILIA_COTA, RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2), RIGHT('0' + LTRIM(YEAR(ITENS.DATA_CEDO)),4) " & Chr(10)
    
    Sql = Sql & "ORDER BY PROD.LINHA_PRODUTO, PROD.FAMILIA_COTA, RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2), RIGHT('0' + LTRIM(YEAR(ITENS.DATA_CEDO)),4) " & Chr(10)
    
    Set RstEstFam = New adodb.Recordset

    RstEstFam.CursorLocation = adUseClient
    RstEstFam.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_EstFamCota = RstEstFam

    Exit Function

ErrorHandler:

    If Not RstEstFam Is Nothing Then
        Set RstEstFam = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_EstFamCota"), Err.Description

End Function

' Descric o : Listar_DescontoMed
' Retorno   : RecordSet
'

Public Function Listar_DescontoMed(ByVal strMESANO As String, _
                           Optional ByVal strLINHA_PRODUTO As String, _
                           Optional ByVal strCATEG_PRODUTO As String, _
                           Optional ByVal strTP_ESTOQUE As String, _
                           Optional ByVal strQUALIDADE As String, _
                           Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim RstDscMed      As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    If Vazio(Trim(strMESANO)) Then
    
        Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Simulador de Desconto M dio"), "MesAno deve ser Informado."
        
    End If

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD.UN_NEGOCIO_PROD, " & Chr(10)
    Sql = Sql & "   PROD.LINHA_PRODUTO, " & Chr(10)
    Sql = Sql & "   SUBSTRING(LINE_F.COD_PRODUTO,17,1) QUALIDADE, " & Chr(10)
    Sql = Sql & "   LINE_F.COT_ESTOQUE_PROD, " & Chr(10)
    Sql = Sql & "   LINE_F.CATEGORIA_PROD, " & Chr(10)
    Sql = Sql & "   SUM(((DSCM.VLR_PRC_UNIT / (1 - (LINE_F.DESCONTO / 100))) - DSCM.VLR_PRC_UNIT)  * DSCM.QTD_SCHEDULED) DESCONTO_CONC, " & Chr(10)
    Sql = Sql & "   SUM(DSCM.VLR_PRC_LISTA * DSCM.QTD_SCHEDULED) FAT_BRUTO, " & Chr(10)
    Sql = Sql & "   SUM(DSCM.QTD_SCHEDULED) VOLUME_TOT " & Chr(10)
    
    Sql = Sql & "   FROM PEDIDO HDR_F, DESC_MEDIO DSCM, ITENS_PEDIDO LINE_F " & Chr(10)
    
    Sql = Sql & " INNER JOIN PRODUTO PROD " & Chr(10)
    Sql = Sql & "       ON(PROD.COD_PRODUTO = LINE_F.COD_PRODUTO " & Chr(10)
    
    If Not Vazio(Trim(strUN_Negocio)) Then
       Sql = Sql & "       AND PROD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "'"
    End If
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
       Sql = Sql & Chr(10)
       Sql = Sql & "       AND PROD.LINHA_PRODUTO = '" & Substitui(strLINHA_PRODUTO) & "'"
    
    End If
    Sql = Sql & ")" & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)

    Sql = Sql & SqlAux & " LINE_F.NUM_PEDIDO = HDR_F.NUM_PEDIDO " & Chr(10)
    
    SqlAux = " AND "
    
    Sql = Sql & SqlAux & " HDR_F.COD_SIT_PEDIDO <> 'C' " & Chr(10)
    Sql = Sql & SqlAux & " LINE_F.SITUACAO <> 'C' " & Chr(10)
    
    If Not Vazio(Trim(strTP_ESTOQUE)) Then
         Sql = Sql & SqlAux & " LINE_F.COT_ESTOQUE_PROD = '" & Substitui(strTP_ESTOQUE) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strQUALIDADE)) Then
         Sql = Sql & SqlAux & " SUBSTRING(LINE_F.COD_PRODUTO,17,1) = '" & Substitui(strQUALIDADE) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strCATEG_PRODUTO)) Then
       Sql = Sql & SqlAux & " LINE_F.CATEGORIA_PROD = '" & Substitui(strCATEG_PRODUTO) & "'"
    End If
    
    If Not Vazio(Trim(strMESANO)) Then
         Sql = Sql & SqlAux & " (MONTH(LINE_F.DATA_REQ_SHIP) = '" & Mid(Substitui(strMESANO), 1, 2) & "'" & Chr(10)
         Sql = Sql & SqlAux & " YEAR(LINE_F.DATA_REQ_SHIP) = '" & Mid(Substitui(strMESANO), 3, 4) & "') " & Chr(10)
    End If

    Sql = Sql & SqlAux & " DSCM.NUM_PEDIDO = LINE_F.NUM_PEDIDO " & Chr(10)
    Sql = Sql & SqlAux & " DSCM.ID_SEQUENCIAL = LINE_F.ID_SEQUENCIAL " & Chr(10)
    Sql = Sql & SqlAux & " DSCM.COD_PRODUTO = LINE_F.COD_PRODUTO " & Chr(10)

    Sql = Sql & "GROUP BY PROD.UN_NEGOCIO_PROD, PROD.LINHA_PRODUTO, SUBSTRING(LINE_F.COD_PRODUTO,17,1), LINE_F.COT_ESTOQUE_PROD, LINE_F.CATEGORIA_PROD" & Chr(10)
    
    Sql = Sql & "ORDER BY PROD.UN_NEGOCIO_PROD, PROD.LINHA_PRODUTO, SUBSTRING(LINE_F.COD_PRODUTO,17,1), LINE_F.COT_ESTOQUE_PROD, LINE_F.CATEGORIA_PROD" & Chr(10)
    
    Set RstDscMed = New adodb.Recordset

    RstDscMed.CursorLocation = adUseClient
    RstDscMed.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_DescontoMed = RstDscMed

    Exit Function

ErrorHandler:

    If Not RstDscMed Is Nothing Then
        Set RstDscMed = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Simulador de Desconto M dio"), Err.Description

End Function

' Descric o : Listar_CatArtPad
' Retorno   : RecordSet
'

Public Function Listar_CatArtPad(ByVal strMESANO As String, _
                           Optional ByVal strLINHA_PRODUTO As String, _
                           Optional ByVal strCATEG_PRODUTO As String, _
                           Optional ByVal strARTPADRAO As String, _
                           Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim RstCatArtPd    As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    If Vazio(Trim(strMESANO)) Then
    
        Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Simulador de Desconto M dio"), "MesAno deve ser Informado."
        
    End If

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD.UN_NEGOCIO_PROD, " & Chr(10)
    Sql = Sql & "   PROD.LINHA_PRODUTO, " & Chr(10)
    Sql = Sql & "   PROD.COD_ARTIGO_PADRAO, " & Chr(10)
    Sql = Sql & "   LINE_F.CATEGORIA_PROD, " & Chr(10)
    Sql = Sql & "   SUM(DSCM.VLR_PRC_UNIT * DSCM.QTD_SCHEDULED) VLR_FATUR, " & Chr(10)
    Sql = Sql & "   SUM(DSCM.QTD_SCHEDULED) VOLUME_TOT " & Chr(10)
    
    Sql = Sql & "   FROM PEDIDO HDR_F, DESC_MEDIO DSCM, ITENS_PEDIDO LINE_F " & Chr(10)
    
    Sql = Sql & " INNER JOIN PRODUTO PROD " & Chr(10)
    Sql = Sql & "       ON(PROD.COD_PRODUTO = LINE_F.COD_PRODUTO " & Chr(10)
    
    If Not Vazio(Trim(strUN_Negocio)) Then
       Sql = Sql & "       AND PROD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "'"
    End If
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
       Sql = Sql & Chr(10)
       Sql = Sql & "       AND PROD.LINHA_PRODUTO = '" & Substitui(strLINHA_PRODUTO) & "'"
    
    End If
    Sql = Sql & ")" & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)

    Sql = Sql & SqlAux & " LINE_F.NUM_PEDIDO = HDR_F.NUM_PEDIDO " & Chr(10)
    
    SqlAux = " AND "
    
    Sql = Sql & SqlAux & " HDR_F.COD_SIT_PEDIDO <> 'C' " & Chr(10)
    Sql = Sql & SqlAux & " LINE_F.SITUACAO <> 'C' " & Chr(10)
    
    If Not Vazio(Trim(strARTPADRAO)) Then
       Sql = Sql & SqlAux & " PROD.COD_ARTIGO_PADRAO = '" & Substitui(strARTPADRAO) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strCATEG_PRODUTO)) Then
       Sql = Sql & SqlAux & " LINE_F.CATEGORIA_PROD = '" & Substitui(strCATEG_PRODUTO) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strMESANO)) Then
         Sql = Sql & SqlAux & " (MONTH(LINE_F.DATA_REQ_SHIP) = '" & Mid(Substitui(strMESANO), 1, 2) & "'" & Chr(10)
         Sql = Sql & SqlAux & " YEAR(LINE_F.DATA_REQ_SHIP) = '" & Mid(Substitui(strMESANO), 3, 4) & "') " & Chr(10)
    End If

    Sql = Sql & SqlAux & " DSCM.NUM_PEDIDO = LINE_F.NUM_PEDIDO " & Chr(10)
    Sql = Sql & SqlAux & " DSCM.ID_SEQUENCIAL = LINE_F.ID_SEQUENCIAL " & Chr(10)
    Sql = Sql & SqlAux & " DSCM.COD_PRODUTO = LINE_F.COD_PRODUTO " & Chr(10)

    Sql = Sql & "GROUP BY PROD.UN_NEGOCIO_PROD, PROD.LINHA_PRODUTO, PROD.COD_ARTIGO_PADRAO, LINE_F.CATEGORIA_PROD" & Chr(10)
    
    Sql = Sql & "ORDER BY PROD.UN_NEGOCIO_PROD, PROD.LINHA_PRODUTO, PROD.COD_ARTIGO_PADRAO, LINE_F.CATEGORIA_PROD" & Chr(10)
    
    Set RstCatArtPd = New adodb.Recordset

    RstCatArtPd.CursorLocation = adUseClient
    RstCatArtPd.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_CatArtPad = RstCatArtPd

    Exit Function

ErrorHandler:

    If Not RstCatArtPd Is Nothing Then
        Set RstCatArtPd = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Consulta por Categoria,Artigo/Padr o"), Err.Description

End Function

Public Function Listar_Vendas_Diarias(ByVal strDATA_INICIAL As String, _
                                      ByVal strDATA_FINAL As String, _
                                      Optional ByVal strCLIENTE As String, _
                                      Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                      Optional ByVal strUN_Negocio As String) As adodb.Recordset


    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim strConnect     As String
    Dim RstVD          As adodb.Recordset
    Dim SqlAux         As String
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Vendas_Diarias"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If
    
    strConnect = ConnectionSQL("IFV")

    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & " CLI.CNPJ," & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.POSTO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " ITENS.DATA_CEDO," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " ITENS.ID_SEQUENCIAL," & Chr(10)
    Sql = Sql & " ITENS.COD_PRODUTO," & Chr(10)
    Sql = Sql & " ITENS.QTD_PEDIDA," & Chr(10)
    Sql = Sql & " PROD.COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " ART.DES_ARTIGO_PADRAO" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    Sql = Sql & " INNER JOIN ITENS_PEDIDO ITENS ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO)" & Chr(10)
    
    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & " RIGHT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO" & Chr(10)
         Sql = Sql & "       AND PROD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "')"
    Else
         Sql = Sql & " RIGHT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO)" & Chr(10)
    End If
    
    Sql = Sql & " LEFT OUTER JOIN ARTIGO_PADRAO ART ON (ART.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO)" & Chr(10)

    Sql = Sql & " WHERE PED.COD_SIT_PEDIDO <> 'C '" & Chr(10)
    Sql = Sql & " AND ITENS.SITUACAO <> 'C '" & Chr(10)
    
    SqlAux = " AND "
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
       Sql = Sql & " AND (ITENS.DATA_CEDO >= " & ToData(strDATA_INICIAL) & Chr(10)
       Sql = Sql & " AND ITENS.DATA_TARDE <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
       
    End If
               
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & "( ITENS.COD_PRODUTO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO) & "%'" & Chr(10)
         Sql = Sql & " OR  ART.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " (CLI.CNPJ LIKE '%" & Substitui(strCLIENTE) & "%'" & Chr(10)
         Sql = Sql & " OR CLI.RAZAO_SOCIAL LIKE '%" & Substitui(strCLIENTE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
               
    Sql = Sql & "ORDER BY CLI.CNPJ, PED.ESTABELECIMENTO, PED.NUM_PEDIDO" & Chr(10)
               
    Set RstVD = New adodb.Recordset
   
    RstVD.CursorLocation = adUseClient
    RstVD.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Vendas_Diarias = RstVD

    Exit Function

ErrorHandler:

    If Not RstVD Is Nothing Then
        Set RstVD = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Vendas_Diarias"), Err.Description

End Function

Public Function Listar_Pedidos_Cancelados(ByVal strDATA_INICIAL As String, _
                                      ByVal strDATA_FINAL As String) As adodb.Recordset


    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim strConnect     As String
    Dim RstPC          As adodb.Recordset
    Dim SqlAux         As String
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Vendas_Diarias"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If
    
    strConnect = ConnectionSQL("IFV")

    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & " PED.CLIENTE," & Chr(10)
    Sql = Sql & " PED.ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    
    SqlAux = " AND "
    
    Sql = Sql & " WHERE PED.COD_SIT_PEDIDO = 'C '" & Chr(10)
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
       Sql = Sql & " AND (PED.DATA_ULT_ALT >= " & ToData(strDATA_INICIAL) & Chr(10)
       Sql = Sql & " AND PED.DATA_ULT_ALT <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
    End If
               
    Sql = Sql & "ORDER BY PED.CLIENTE, PED.ESTABELECIMENTO, PED.NUM_PEDIDO" & Chr(10)
               
    Set RstPC = New adodb.Recordset
   
    RstPC.CursorLocation = adUseClient
    RstPC.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Pedidos_Cancelados = RstPC

    Exit Function

ErrorHandler:

    If Not RstPC Is Nothing Then
        Set RstPC = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Pedidos_Cancelados"), Err.Description

End Function

Public Function Listar_PedNaoLiquidados(ByVal strDATA_INICIAL As String, _
                                      ByVal strDATA_FINAL As String, _
                                      Optional ByVal strCLIENTE As String, _
                                      Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                      Optional ByVal strUN_Negocio As String) As adodb.Recordset


    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim strConnect     As String
    Dim RstPL          As adodb.Recordset
    Dim SqlAux         As String
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_Vendas_Diarias"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If
    
    strConnect = ConnectionSQL("IFV")

    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & " CLI.CNPJ," & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.POSTO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " ITENS.DATA_CEDO," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " ITENS.ID_SEQUENCIAL," & Chr(10)
    Sql = Sql & " ITENS.COD_PRODUTO," & Chr(10)
    Sql = Sql & " ITENS.QTD_PEDIDA," & Chr(10)
    Sql = Sql & " PROD.COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " ART.DES_ARTIGO_PADRAO" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    Sql = Sql & " INNER JOIN ITENS_PEDIDO ITENS ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO)" & Chr(10)
    
    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & " RIGHT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO" & Chr(10)
         Sql = Sql & "       AND PROD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "')"
    Else
         Sql = Sql & " RIGHT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO)" & Chr(10)
    End If
    
    Sql = Sql & " LEFT OUTER JOIN ARTIGO_PADRAO ART ON (ART.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO)" & Chr(10)

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
       Sql = Sql & " WHERE (ITENS.DATA_CEDO >= " & ToData(strDATA_INICIAL) & Chr(10)
       Sql = Sql & " AND ITENS.DATA_TARDE <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
       SqlAux = " AND "
    
    End If
    
    Sql = Sql & SqlAux & " ITENS.SITUACAO NOT IN ('F ','C ')" & Chr(10)
    
    SqlAux = " AND "
    
    Sql = Sql & SqlAux & " PED.COD_SIT_PEDIDO NOT IN ('F ','C ')" & Chr(10)
               
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & "( ITENS.COD_PRODUTO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO) & "%'" & Chr(10)
         Sql = Sql & " OR  ART.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " (CLI.CNPJ LIKE '%" & Substitui(strCLIENTE) & "%'" & Chr(10)
         Sql = Sql & " OR CLI.RAZAO_SOCIAL LIKE '%" & Substitui(strCLIENTE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
               
    Sql = Sql & "ORDER BY CLI.CNPJ, PED.ESTABELECIMENTO, PED.NUM_PEDIDO" & Chr(10)
               
    Set RstPL = New adodb.Recordset
   
    RstPL.CursorLocation = adUseClient
    RstPL.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_PedNaoLiquidados = RstPL

    Exit Function

ErrorHandler:

    If Not RstPL Is Nothing Then
        Set RstPL = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_PedNaoLiquidados"), Err.Description

End Function

Public Function Listar_PedPendenciaFinCom(ByVal strDATA_INICIAL As String, _
                                      ByVal strDATA_FINAL As String, _
                                      Optional ByVal strCLIENTE As String) As adodb.Recordset


    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim strConnect     As String
    Dim RstPF          As adodb.Recordset
    Dim SqlAux         As String
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
        If DateValue(strDATA_INICIAL) > DateValue(strDATA_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Listar_PedPendenciaFinCom"), "Data Fim deve conter valor maior ou igual Data In cio."
        End If
        
    End If
    
    strConnect = ConnectionSQL("IFV")

    Sql = Sql & " SELECT DISTINCT" & Chr(10)
    Sql = Sql & " CLI.CNPJ," & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " BLOQ.NUM_LINHA," & Chr(10)
    Sql = Sql & " BLOQ.ID_SEQUENCIAL," & Chr(10)
    Sql = Sql & " ISNULL((SELECT ITENS.COD_PRODUTO FROM ITENS_PEDIDO ITENS WHERE ITENS.NUM_PEDIDO = BLOQ.NUM_PEDIDO AND    ITENS.ID_SEQUENCIAL = BLOQ.NUM_LINHA),'') as COD_PRODUTO," & Chr(10)
    Sql = Sql & " ISNULL((SELECT ITENS.VLR_PRECO FROM ITENS_PEDIDO ITENS WHERE ITENS.NUM_PEDIDO = BLOQ.NUM_PEDIDO AND    ITENS.ID_SEQUENCIAL = BLOQ.NUM_LINHA),0) AS VLR_PRECO," & Chr(10)
    Sql = Sql & " BLOQ.DES_MENSAGEM," & Chr(10)
    Sql = Sql & " BLOQ.DES_BLOQUEIO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL" & Chr(10)
    Sql = Sql & " FROM PEDIDO_BLOQUEIOS BLOQ " & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN PEDIDO PED ON (PED.NUM_PEDIDO = BLOQ.NUM_PEDIDO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)

    Sql = Sql & " WHERE BLOQ.STATUS = 'A'" & Chr(10)
    
    If Not Vazio(Trim(strDATA_INICIAL)) And Not Vazio(Trim(strDATA_FINAL)) Then
    
       Sql = Sql & " AND (PED.DATA_REQUERIDA >= " & ToData(strDATA_INICIAL) & Chr(10)
       Sql = Sql & " AND PED.DATA_REQUERIDA <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
       
    End If
    
    SqlAux = " AND "
               
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " (CLI.CNPJ LIKE '%" & Substitui(strCLIENTE) & "%'" & Chr(10)
         Sql = Sql & " OR CLI.RAZAO_SOCIAL LIKE '%" & Substitui(strCLIENTE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
               
    Sql = Sql & "ORDER BY CLI.CNPJ, PED.NUM_PEDIDO" & Chr(10)
               
    Set RstPF = New adodb.Recordset
   
    RstPF.CursorLocation = adUseClient
    RstPF.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_PedPendenciaFinCom = RstPF

    Exit Function

ErrorHandler:

    If Not RstPF Is Nothing Then
        Set RstPF = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_PedPendenciaFinCom"), Err.Description

End Function



## IFV01_PREVISAO_COTA_VENDA.bas
Attribute VB_Name = "IFV01_PREVISAO_COTA_VENDA"
Private Const strNomeModulo = "IFV01_PREVISAO_COTA_VENDA"

'
' Descric o : Listar_Previsao_Cota_Venda
' Retorno   : RecordSet
'

Public Function Listar_Previsao_Cota_Venda(Optional ByVal strLIKE_FAMILIA As String, _
                                           Optional ByVal strLIKE_MESANO As String, _
                                           Optional ByVal strLINHA_PRODUTO As String) As adodb.Recordset
                                          
                                          

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PREV_CT_V.MESANO , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.LINHA_PRODUTO , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.FAMILIA , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.CATEGORIA , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.FABRICA , " & Chr(10)
    Sql = Sql & "   ISNULL(PREV_CT_V.PREVISAO,'0') PREVISAO , " & Chr(10)
    Sql = Sql & "   ISNULL(PREV_CT_V.COTA,'0')  COTA, " & Chr(10)
    Sql = Sql & "   ISNULL(PREV_CT_V.VENDA,'0')  VENDA " & Chr(10)
                
    Sql = Sql & "   FROM PREVISAO_COTA_VENDA PREV_CT_V " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strLIKE_FAMILIA)) Then
         Sql = Sql & SqlAux & " PREV_CT_V.FAMILIA  LIKE '%" & Substitui(strLIKE_FAMILIA) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_MESANO)) Then
         Sql = Sql & SqlAux & " PREV_CT_V.MESANO  LIKE '%" & Substitui(strLIKE_MESANO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PREV_CT_V.LINHA_PRODUTO LIKE '" & Substitui(strLINHA_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
               
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Previsao_Cota_Venda = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Previsao_Cota_Venda"), Err.Description

End Function
'
' Descric o : Listar_Previsao_Cota_Venda
' Retorno   : RecordSet
'

Public Function Listar_Cota_Cota_Venda(Optional ByVal strLIKE_FAMILIA As String, _
                                       Optional ByVal strLIKE_MESANO As String, _
                                       Optional ByVal strLINHA_PRODUTO As String) As adodb.Recordset
                                          
                                          

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   COT_V.MESANO , " & Chr(10)
    Sql = Sql & "   COT_V.LINHA_PRODUTO , " & Chr(10)
    Sql = Sql & "   COT_V.FAMILIA , " & Chr(10)
    Sql = Sql & "   COT_V.REGIONAL , " & Chr(10)
    Sql = Sql & "   COT_V.CATEGORIA , " & Chr(10)
    Sql = Sql & "   COT_V.FABRICA , " & Chr(10)
    Sql = Sql & "   COT_V.CICLO , " & Chr(10)
    Sql = Sql & "   ISNULL(COT_V.PREVISAO,'0') PREVISAO , " & Chr(10)
    Sql = Sql & "   ISNULL(COT_V.COTA,'0')  COTA, " & Chr(10)
    Sql = Sql & "   ISNULL(COT_V.VENDA,'0')  VENDA " & Chr(10)
                
    Sql = Sql & "   FROM COTA_COTA_VENDA COT_V " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strLIKE_FAMILIA)) Then
         Sql = Sql & SqlAux & " COT_V.FAMILIA  LIKE '%" & Substitui(strLIKE_FAMILIA) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_MESANO)) Then
         Sql = Sql & SqlAux & " COT_V.MESANO  LIKE '%" & Substitui(strLIKE_MESANO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " COT_V.LINHA_PRODUTO LIKE '" & Substitui(strLINHA_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
               
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Cota_Cota_Venda = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Cota_Cota_Venda"), Err.Description

End Function



## IFV01_PRODUTO.bas
Attribute VB_Name = "IFV01_PRODUTO"
Private Const strNomeModulo = "IFV01_PRODUTO"

'
' Descric o : Listar_Produto
' Retorno   : RecordSet
'

Public Function Listar_Produto(Optional ByVal strCOD_PRODUTO As String, _
                               Optional ByVal strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD.COD_PRODUTO , " & Chr(10)
    Sql = Sql & "   PROD.COD_ARTIGO_PADRAO , " & Chr(10)
    Sql = Sql & "   PROD.COD_DIMENSAO , " & Chr(10)
    Sql = Sql & "   PROD.COD_COR , " & Chr(10)
    Sql = Sql & "   PROD.DES_PRODUTO_CURTA , " & Chr(10)
    Sql = Sql & "   PROD.DES_PRODUTO_LONGA , " & Chr(10)
    Sql = Sql & "   PROD.IND_ATIVO , " & Chr(10)
    Sql = Sql & "   PROD.DES_QUALIDADE , " & Chr(10)
    Sql = Sql & "   PROD.COD_EMBALAGEM , " & Chr(10)
    Sql = Sql & "   PROD.DES_TIPO_PRODUTO , " & Chr(10)
    Sql = Sql & "   PROD.COD_QUALIDADE , " & Chr(10)
    Sql = Sql & "   ART_PAD.DES_ARTIGO_PADRAO , " & Chr(10)
    Sql = Sql & "   DIMEN.DES_DIMENSAO , " & Chr(10)
    Sql = Sql & "   COR.DES_COR , " & Chr(10)
    Sql = Sql & "   COR.GRUPOCOR , " & Chr(10)
    Sql = Sql & "   EMB.DES_EMBALAGEM , " & Chr(10)
    Sql = Sql & "   EMB.QTD_METROS  " & Chr(10)
    
        
    Sql = Sql & "   FROM PRODUTO PROD " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN ARTIGO_PADRAO ART_PAD " & Chr(10)
    Sql = Sql & "        ON (ART_PAD.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN DIMENSAO DIMEN " & Chr(10)
    Sql = Sql & "        ON (DIMEN.COD_DIMENSAO = PROD.COD_DIMENSAO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN COR COR " & Chr(10)
    Sql = Sql & "        ON (COR.COD_COR = PROD.COD_COR) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN EMBALAGEM EMB " & Chr(10)
    Sql = Sql & "        ON (EMB.COD_EMBALAGEM = PROD.COD_EMBALAGEM) " & Chr(10)
   
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.COD_PRODUTO = '" & Substitui(strCOD_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA)) Then
         Sql = Sql & SqlAux & "( PROD.COD_PRODUTO LIKE '%" & Substitui(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_PRODUTO_CURTA LIKE '%" & Substitui(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY PROD.COD_PRODUTO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Produto = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Produto"), Err.Description

End Function

'
' Descric o : Listar_Artigo_Padrao
' Retorno   : RecordSet
'

Public Function Listar_Artigo_Padrao(Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                     Optional ByVal strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT DISTINCT " & Chr(10)
    Sql = Sql & "   ART_PAD.COD_ARTIGO_PADRAO , " & Chr(10)
    Sql = Sql & "   ART_PAD.DES_ARTIGO_PADRAO,  " & Chr(10)
    Sql = Sql & "   PROD.CATEGORIA_PROD  " & Chr(10)
            
    Sql = Sql & "   FROM ARTIGO_PADRAO ART_PAD " & Chr(10)
    Sql = Sql & "   INNER JOIN PRODUTO PROD " & Chr(10)
    Sql = Sql & "     ON (PROD.COD_ARTIGO_PADRAO = ART_PAD.COD_ARTIGO_PADRAO) " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " ART_PAD.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & "( ART_PAD.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%'" & Chr(10)
         Sql = Sql & " OR  ART_PAD.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY ART_PAD.COD_ARTIGO_PADRAO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Artigo_Padrao = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Artigo_Padrao"), Err.Description

End Function


'
' Descric o : Listar_Cor
' Retorno   : RecordSet
'

Public Function Listar_Cor(Optional ByVal strCOD_COR As String, _
                           Optional ByVal strLIKE_COD_COR_DES_COR As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   COR.COD_COR , " & Chr(10)
    Sql = Sql & "   COR.DES_COR , " & Chr(10)
    Sql = Sql & "   COR.GRUPOCOR , " & Chr(10)
    Sql = Sql & "   COR.IND_ATIVO  " & Chr(10)
            
    Sql = Sql & "   FROM COR COR " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_COR)) Then
         Sql = Sql & SqlAux & " COR.COD_COR = '" & Substitui(strCOD_COR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_COR_DES_COR)) Then
         Sql = Sql & SqlAux & "( COR.COD_COR LIKE '%" & Substitui(strLIKE_COD_COR_DES_COR) & "%'" & Chr(10)
         Sql = Sql & " OR  COR.DES_COR LIKE '%" & Substitui(strLIKE_COD_COR_DES_COR) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY COR.COD_COR " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Cor = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Cor"), Err.Description

End Function

'
' Descric o : Listar_Fornecedor
' Retorno   : RecordSet
'

Public Function Listar_Fornecedor(Optional ByVal strCOD_FORNECEDOR As String, _
                                  Optional ByVal strLIKE_COD_FORNECEDOR_DES_FORNECEDOR As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   FORN.COD_FORNECEDOR , " & Chr(10)
    Sql = Sql & "   FORN.DES_FORNECEDOR , " & Chr(10)
    Sql = Sql & "   FORN.STATUS  " & Chr(10)
    
            
    Sql = Sql & "   FROM FORNECEDOR FORN " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_FORNECEDOR)) Then
         Sql = Sql & SqlAux & " FORN.COD_FORNECEDOR = '" & Substitui(strCOD_FORNECEDOR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_FORNECEDOR_DES_FORNECEDOR)) Then
         Sql = Sql & SqlAux & "( FORN.COD_FORNECEDOR LIKE '%" & Substitui(strLIKE_COD_FORNECEDOR_DES_FORNECEDOR) & "%'" & Chr(10)
         Sql = Sql & " OR  FORN.DES_FORNECEDOR LIKE '%" & Substitui(strLIKE_COD_FORNECEDOR_DES_FORNECEDOR) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY FORN.DES_FORNECEDOR " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Fornecedor = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Fornecedor"), Err.Description

End Function

'
' Descric o : Listar_Produto_Concorrente_Concorrente
' Retorno   : RecordSet
'

Public Function Listar_Produto_Concorrente(Optional ByVal strCOD_FORNECEDOR As String, _
                                           Optional ByVal strLIKE_COD_FORNECEDOR_DES_FORNECEDOR As String, _
                                           Optional ByVal strSIMILAR_AST As String, _
                                           Optional ByVal strLIKE_SIMILAR_AST_DES_PRODUTO_CURTA As String, _
                                           Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                           Optional ByVal strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO As String, _
                                           Optional ByVal strCOD_COR As String, _
                                           Optional ByVal strLIKE_COD_COR_DES_COR As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD_CONC.COD_FORNECEDOR , " & Chr(10)
    Sql = Sql & "   PROD_CONC.COD_PRODUTO , " & Chr(10)
    Sql = Sql & "   PROD_CONC.SIMILAR_AST , " & Chr(10)
    Sql = Sql & "   PROD_CONC.NOME_PRODUTO , " & Chr(10)
    Sql = Sql & "   PROD_CONC.COD_PRODUTO_FORNECEDOR , " & Chr(10)
    Sql = Sql & "   PROD_CONC.LARGURA , " & Chr(10)
    Sql = Sql & "   PROD_CONC.PESO , " & Chr(10)
    Sql = Sql & "   PROD.COD_ARTIGO_PADRAO, " & Chr(10)
    Sql = Sql & "   FORN.DES_FORNECEDOR , " & Chr(10)
    Sql = Sql & "   PROD.DES_PRODUTO_CURTA , " & Chr(10)
    Sql = Sql & "   PROD.DES_PRODUTO_LONGA , " & Chr(10)
    Sql = Sql & "   ART_PAD.DES_ARTIGO_PADRAO , " & Chr(10)
    Sql = Sql & "   DIMEN.DES_DIMENSAO , " & Chr(10)
    Sql = Sql & "   COR.DES_COR , " & Chr(10)
    Sql = Sql & "   COR.GRUPOCOR , " & Chr(10)
    Sql = Sql & "   EMB.DES_EMBALAGEM , " & Chr(10)
    Sql = Sql & "   EMB.QTD_METROS  " & Chr(10)
            
    Sql = Sql & "   FROM PRODUTO_CONCORRENTE  PROD_CONC " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN FORNECEDOR FORN " & Chr(10)
    Sql = Sql & "        ON (FORN.COD_FORNECEDOR = PROD_CONC.COD_FORNECEDOR) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN PRODUTO PROD " & Chr(10)
    Sql = Sql & "        ON (PROD.COD_PRODUTO = PROD_CONC.SIMILAR_AST) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN ARTIGO_PADRAO ART_PAD " & Chr(10)
    Sql = Sql & "        ON (ART_PAD.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN DIMENSAO DIMEN " & Chr(10)
    Sql = Sql & "        ON (DIMEN.COD_DIMENSAO = PROD.COD_DIMENSAO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN COR COR " & Chr(10)
    Sql = Sql & "        ON (COR.COD_COR = PROD.COD_COR) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN EMBALAGEM EMB " & Chr(10)
    Sql = Sql & "        ON (EMB.COD_EMBALAGEM = PROD.COD_EMBALAGEM) " & Chr(10)
    
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_FORNECEDOR)) Then
         Sql = Sql & SqlAux & " PROD_CONC.COD_FORNECEDOR = '" & Substitui(strCOD_FORNECEDOR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_FORNECEDOR_DES_FORNECEDOR)) Then
         Sql = Sql & SqlAux & "( PROD_CONC.COD_FORNECEDOR LIKE '%" & Substitui(strLIKE_COD_FORNECEDOR_DES_FORNECEDOR) & "%'" & Chr(10)
         Sql = Sql & " OR  FORN.DES_FORNECEDOR LIKE '%" & Substitui(strLIKE_COD_PRODUTO_DES_PRODUTO_CURTA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
     
    If Not Vazio(Trim(strSIMILAR_AST)) Then
         Sql = Sql & SqlAux & " PROD_CONC.SIMILAR_AST = '" & Substitui(strSIMILAR_AST) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_SIMILAR_AST_DES_PRODUTO_CURTA)) Then
         Sql = Sql & SqlAux & "( PROD_CONC.SIMILAR_AST LIKE '%" & Substitui(strLIKE_SIMILAR_AST_DES_PRODUTO_CURTA) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_PRODUTO_CURTA LIKE '%" & Substitui(strLIKE_SIMILAR_AST_DES_PRODUTO_CURTA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " ART_PAD.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & "( ART_PAD.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%'" & Chr(10)
         Sql = Sql & " OR  ART_PAD.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_COR)) Then
         Sql = Sql & SqlAux & " COR.COD_COR = '" & Substitui(strCOD_COR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_COR_DES_COR)) Then
         Sql = Sql & SqlAux & "( COR.COD_COR LIKE '%" & Substitui(strLIKE_COD_COR_DES_COR) & "%'" & Chr(10)
         Sql = Sql & " OR  COR.DES_COR LIKE '%" & Substitui(strLIKE_COD_COR_DES_COR) & "%')" & Chr(10)
         SqlAux = " AND "
    End If

    Sql = Sql & "ORDER BY PROD_CONC.SIMILAR_AST " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Produto_Concorrente = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Produto_Concorrente"), Err.Description

End Function


'
' Descric o : Listar_Qualidade
' Retorno   : RecordSet
'

Public Function Listar_Qualidade(Optional ByVal strCOD_QUALIDADE As String, _
                                 Optional ByVal strLIKE_COD_QUALIDADE_DES_QUALIDADE As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   DISTINCT QUAL.COD_QUALIDADE , " & Chr(10)
    Sql = Sql & "   QUAL.DES_QUALIDADE  " & Chr(10)
    
    Sql = Sql & "   FROM PRODUTO QUAL " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_QUALIDADE)) Then
         Sql = Sql & SqlAux & " QUAL.COD_QUALIDADE = '" & Substitui(strCOD_QUALIDADE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_QUALIDADE_DES_QUALIDADE)) Then
         Sql = Sql & SqlAux & "( QUAL.COD_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%'" & Chr(10)
         Sql = Sql & " OR  QUAL.DES_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY QUAL.DES_QUALIDADE " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Qualidade = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Qualidade"), Err.Description

End Function

'
' Descric o : Listar_Embalagem
' Retorno   : RecordSet
'

Public Function Listar_Embalagem(Optional ByVal strCOD_EMBALAGEM As String, _
                                     Optional ByVal strLIKE_COD_EMBALAGEM_DES_EMBALAGEM As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   EMB.COD_EMBALAGEM , " & Chr(10)
    Sql = Sql & "   EMB.DES_EMBALAGEM  " & Chr(10)
            
    Sql = Sql & "   FROM EMBALAGEM EMB " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_EMBALAGEM)) Then
         Sql = Sql & SqlAux & " EMB.COD_EMBALAGEM = '" & Substitui(strCOD_EMBALAGEM) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_EMBALAGEM_DES_EMBALAGEM)) Then
         Sql = Sql & SqlAux & "( EMB.COD_EMBALAGEM LIKE '%" & Substitui(strLIKE_COD_EMBALAGEM_DES_EMBALAGEM) & "%'" & Chr(10)
         Sql = Sql & " OR  EMB.DES_EMBALAGEM LIKE '%" & Substitui(strLIKE_COD_EMBALAGEM_DES_EMBALAGEM) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    Sql = Sql & "ORDER BY EMB.COD_EMBALAGEM " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Embalagem = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Embalagem"), Err.Description

End Function

'
' Descric o : Listar_Linha
' Retorno   : RecordSet
'

Public Function Listar_Linha(Optional ByVal strLINHA_PRODUTO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   DISTINCT PROD.LINHA_PRODUTO " & Chr(10)
    Sql = Sql & "   FROM PRODUTO PROD " & Chr(10)
       
    SqlAux = " WHERE  " & Chr(10)
        
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.LINHA_PRODUTO = '" & Substitui(strLINHA_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY PROD.LINHA_PRODUTO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Linha = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Linha"), Err.Description

End Function


'
' Descric o : Listar_Linha_Agrupado
' Retorno   : RecordSet
'

Public Function Listar_Linha_Agrupado(ByVal strANO As String, _
                                      ByVal strCLIENTE As String, _
                                      Optional ByVal strLINHA_PRODUTO As String, _
                                      Optional ByVal strCOD_QUALIDADE As String, _
                                      Optional ByVal strLIKE_COD_QUALIDADE_DES_QUALIDADE As String, _
                                      Optional ByVal strAGRUPADO_MES As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD.LINHA_PRODUTO, " & Chr(10)
    If Not Vazio(strAGRUPADO_MES) Then
        Sql = Sql & "   RIGHT('0' + LTRIM(MONTH(PED.DATA_REQUERIDA)),2) DATA_REQUERIDA, " & Chr(10)
    End If
    Sql = Sql & "   ISNULL(SUM(ITENS_PED.QTD_FATURADA),'0') QTD_FATURADA " & Chr(10)
       
    Sql = Sql & "   FROM PRODUTO PROD " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN ITENS_PEDIDO ITENS_PED " & Chr(10)
    Sql = Sql & "   ON(ITENS_PED.COD_PRODUTO = PROD.COD_PRODUTO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN PEDIDO PED " & Chr(10)
    Sql = Sql & "   ON(PED.NUM_PEDIDO = ITENS_PED.NUM_PEDIDO) " & Chr(10)
    Sql = Sql & "   Where ITENS_PED.QTD_FATURADA Is Not Null " & Chr(10)
    
    SqlAux = " AND  " & Chr(10)
    
    If Not Vazio(Trim(strANO)) Then
         Sql = Sql & SqlAux & " YEAR(PED.DATA_REQUERIDA) = " & strANO & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " PED.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.LINHA_PRODUTO LIKE '%" & Substitui(strLINHA_PRODUTO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_QUALIDADE)) Then
         Sql = Sql & SqlAux & " PROD.COD_QUALIDADE = '" & Substitui(strCOD_QUALIDADE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_QUALIDADE_DES_QUALIDADE)) Then
         Sql = Sql & SqlAux & "( PROD.COD_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Vazio(strAGRUPADO_MES) Then
        Sql = Sql & "   GROUP BY LINHA_PRODUTO "
        Sql = Sql & "   ORDER BY PROD.LINHA_PRODUTO " & Chr(10)
    Else
        Sql = Sql & "   GROUP BY LINHA_PRODUTO,RIGHT('0' + LTRIM(MONTH(PED.DATA_REQUERIDA)),2) "
        Sql = Sql & "   ORDER BY RIGHT('0' + LTRIM(MONTH(PED.DATA_REQUERIDA)),2) " & Chr(10)
    End If
    
    
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Linha_Agrupado = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Linha_Agrupado"), Err.Description

End Function



'
' Descric o : Listar_Aritgo_Padrao_Agrupado
' Retorno   : RecordSet
'

Public Function Listar_Aritgo_Padrao_Agrupado(ByVal strANO As String, _
                                              ByVal strCLIENTE As String, _
                                              Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                              Optional ByVal strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO As String, _
                                              Optional ByVal strCOD_QUALIDADE As String, _
                                              Optional ByVal strLIKE_COD_QUALIDADE_DES_QUALIDADE As String, _
                                              Optional ByVal strAGRUPADO_MES As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   PROD.COD_ARTIGO_PADRAO, " & Chr(10)
    If Not Vazio(strAGRUPADO_MES) Then
        Sql = Sql & "   RIGHT('0' + LTRIM(MONTH(PED.DATA_REQUERIDA)),2) DATA_REQUERIDA, " & Chr(10)
    End If
    Sql = Sql & "   ISNULL(SUM(ITENS_PED.QTD_FATURADA),'0') QTD_FATURADA " & Chr(10)
    Sql = Sql & "   FROM PRODUTO PROD " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN ITENS_PEDIDO ITENS_PED " & Chr(10)
    Sql = Sql & "   ON(ITENS_PED.COD_PRODUTO = PROD.COD_PRODUTO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN ARTIGO_PADRAO ART_PAD " & Chr(10)
    Sql = Sql & "   ON(ART_PAD.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO) " & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN PEDIDO PED " & Chr(10)
    Sql = Sql & "   ON(PED.NUM_PEDIDO = ITENS_PED.NUM_PEDIDO) " & Chr(10)
    Sql = Sql & "   Where ITENS_PED.QTD_FATURADA Is Not Null " & Chr(10)
    
    SqlAux = " AND  " & Chr(10)
    
    If Not Vazio(Trim(strANO)) Then
         Sql = Sql & SqlAux & " YEAR(PED.DATA_REQUERIDA) = " & strANO & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " PED.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " PROD.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & "( PROD.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%'" & Chr(10)
         Sql = Sql & " OR  ART_PAD.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strLIKE_COD_ARTIGO_PADRAO_DES_ARTIGO_PADRAO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_QUALIDADE)) Then
         Sql = Sql & SqlAux & " PROD.COD_QUALIDADE = '" & Substitui(strCOD_QUALIDADE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_COD_QUALIDADE_DES_QUALIDADE)) Then
         Sql = Sql & SqlAux & "( PROD.COD_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_QUALIDADE LIKE '%" & Substitui(strLIKE_COD_QUALIDADE_DES_QUALIDADE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Vazio(strAGRUPADO_MES) Then
        Sql = Sql & "   GROUP BY PROD.COD_ARTIGO_PADRAO "
        Sql = Sql & "ORDER BY PROD.COD_ARTIGO_PADRAO " & Chr(10)
    Else
        Sql = Sql & "   GROUP BY PROD.COD_ARTIGO_PADRAO,RIGHT('0' + LTRIM(MONTH(PED.DATA_REQUERIDA)),2) "
        Sql = Sql & "ORDER BY RIGHT('0' + LTRIM(MONTH(PED.DATA_REQUERIDA)),2) " & Chr(10)
    End If
    
    
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Aritgo_Padrao_Agrupado = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Aritgo_Padrao_Agrupado"), Err.Description

End Function



'
' Descric o : Listar_Data
' Retorno   : RecordSet
'

Public Function Listar_Data() As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   DISTINCT YEAR(PED.DATA_REQUERIDA) DATA_REQUERIDA " & Chr(10)
    Sql = Sql & "   FROM PEDIDO PED " & Chr(10)
    Sql = Sql & " ORDER BY YEAR(PED.DATA_REQUERIDA) " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Data = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Data"), Err.Description

End Function

'
' Descric o : Listar_Categoria
' Retorno   : RecordSet
'

Public Function Listar_Categoria(Optional ByVal strCATEG_PRODUTO As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   DISTINCT PROD.CATEGORIA_PROD " & Chr(10)
    Sql = Sql & "   FROM PRODUTO PROD " & Chr(10)
       
    SqlAux = " WHERE  " & Chr(10)
        
    If Not Vazio(Trim(strCATEG_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.CATEGORIA_PROD = '" & Substitui(strCATEG_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY PROD.CATEGORIA_PROD " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Categoria = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Categoria"), Err.Description

End Function

'
' Descric o : Listar_Familia_Cota
' Retorno   : RecordSet
'

Public Function Listar_Familia_Cota(Optional ByVal strFamilia_COTA As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")

    Sql = " SELECT " & Chr(10)
    Sql = Sql & "   DISTINCT COTA.FAMILIA " & Chr(10)
    Sql = Sql & "   FROM COTA_COTA_VENDA COTA " & Chr(10)
       
    SqlAux = " WHERE  " & Chr(10)
        
    If Not Vazio(Trim(strFamilia_COTA)) Then
         Sql = Sql & SqlAux & " COTA.FAMILIA = '" & Substitui(strFamilia_COTA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY COTA.FAMILIA " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Familia_Cota = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Categoria"), Err.Description

End Function

'
' Descric o : Listar_Produto_Tela
' Retorno   : RecordSet
'

Public Function Listar_Produtos_Tela(Optional ByVal strCOD_PRODUTO As String, _
                                     Optional ByVal strStatus As String, _
                                     Optional ByVal strUN_Negocio As String) As adodb.Recordset

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Rst            As adodb.Recordset
    Dim strConnect     As String
    Dim SqlAux         As String
    
    strConnect = ConnectionSQL("IFV")
    
    Sql = " SELECT DISTINCT " & Chr(10)
    Sql = Sql & "   PROD.COD_PRODUTO , " & Chr(10)
    Sql = Sql & "   PROD.COD_ARTIGO_PADRAO , " & Chr(10)
    Sql = Sql & "   PROD.COD_DIMENSAO , " & Chr(10)
    Sql = Sql & "   PROD.COD_COR , " & Chr(10)
    Sql = Sql & "   PROD.COD_EMBALAGEM , " & Chr(10)
    Sql = Sql & "   PROD.COD_QUALIDADE , " & Chr(10)
    Sql = Sql & "   PROD.IND_ATIVO , " & Chr(10)
    Sql = Sql & "   PROD.UN_NEGOCIO_PROD , " & Chr(10)
    Sql = Sql & "   PROD.CATEGORIA_PROD, " & Chr(10)
    Sql = Sql & "   PROD.FAMILIA_COTA, " & Chr(10)
    Sql = Sql & "   PROD.DES_PRODUTO_LONGA, " & Chr(10)
    Sql = Sql & "   ITENS_NF.POSFISC " & Chr(10)
    Sql = Sql & "   FROM PRODUTO PROD, ITENS_NOTA_FISCAL ITENS_NF " & Chr(10)
    SqlAux = " WHERE " & Chr(10)
    
    Sql = Sql & SqlAux & " PROD.COD_PRODUTO = ITENS_NF.COD_PRODUTO " & Chr(10)
    SqlAux = " AND "
    
    If Not Vazio(Trim(strCOD_PRODUTO)) Then
         Sql = Sql & SqlAux & "( PROD.COD_PRODUTO LIKE '%" & Substitui(strCOD_PRODUTO) & "%'" & Chr(10)
         Sql = Sql & " OR  PROD.DES_PRODUTO_CURTA LIKE '%" & Substitui(strCOD_PRODUTO) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strStatus)) Then
         Sql = Sql & SqlAux & "PROD.IND_ATIVO = '" & Substitui(strStatus) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & SqlAux & " PROD.UN_NEGOCIO_PROD = '" & Substitui(strUN_Negocio) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY PROD.COD_PRODUTO " & Chr(10)
           
    Set Rst = New adodb.Recordset

    Rst.CursorLocation = adUseClient
    Rst.Open Sql, strConnect, adOpenForwardOnly, adLockReadOnly

    Set Listar_Produtos_Tela = Rst

    Exit Function

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Listar_Produtos_Tela"), Err.Description

End Function



## IFV01_RELATORIOS.bas
Attribute VB_Name = "IFV01_RELATORIOS"
Private Const strNomeModulo = "IFV01_RELATORIOS"

Public Sub Relatorio_Cliente(ByVal strCOMPUTADOR As String, _
                             Optional ByVal strCLIENTE As String, _
                             Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                             Optional ByVal strUF As String, _
                             Optional ByVal strLIKE_CIDADE As String, _
                             Optional ByVal strStatus As String, _
                             Optional ByVal strNIVEL_CLIENTE As String, _
                             Optional ByVal strAtendido As String, _
                             Optional ByVal strUN_Negocio As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_CLIENTE WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
        
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_CLIENTE"
        
    Sql = " INSERT INTO DB_IFV.dbo.TMP_CLIENTE " & Chr(10)
    Sql = Sql & "(COMPUTADOR, " & Chr(10)
    Sql = Sql & "CLIENTE , " & Chr(10)
    Sql = Sql & "CNPJ , " & Chr(10)
    Sql = Sql & "ESTABELECIMENTO , " & Chr(10)
    Sql = Sql & "RAZAO_SOCIAL , " & Chr(10)
    Sql = Sql & "NOME_FANTASIA , " & Chr(10)
    Sql = Sql & "DATA_CADASTRAMENTO , " & Chr(10)
    Sql = Sql & "LOGRADOURO , " & Chr(10)
    Sql = Sql & "NUM_LOGRADOURO , " & Chr(10)
    Sql = Sql & "COMPL_LOGRADOURO , " & Chr(10)
    Sql = Sql & "BAIRRO , " & Chr(10)
    Sql = Sql & "CIDADE , " & Chr(10)
    Sql = Sql & "UF , " & Chr(10)
    Sql = Sql & "CEP , " & Chr(10)
    Sql = Sql & "NIVEL_CLIENTE , " & Chr(10)
    Sql = Sql & "INSCRICAO_ESTADUAL , " & Chr(10)
    Sql = Sql & "COBRANCA , " & Chr(10)
    Sql = Sql & "ENTREGA , " & Chr(10)
    Sql = Sql & "COMERCIAL,  " & Chr(10)
    Sql = Sql & "STATUS)  " & Chr(10)
            
    Sql = Sql & " SELECT DISTINCT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & "   CLIENTE.CLIENTE , " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ , " & Chr(10)
    Sql = Sql & "   CLIENTE.ESTABELECIMENTO , " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL , " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA , " & Chr(10)
    Sql = Sql & "   CLIENTE.DATA_CADASTRAMENTO , " & Chr(10)
    Sql = Sql & "   CLIENTE.LOGRADOURO , " & Chr(10)
    Sql = Sql & "   CLIENTE.NUM_LOGRADOURO , " & Chr(10)
    Sql = Sql & "   CLIENTE.COMPL_LOGRADOURO , " & Chr(10)
    Sql = Sql & "   CLIENTE.BAIRRO , " & Chr(10)
    Sql = Sql & "   CLIENTE.CIDADE , " & Chr(10)
    Sql = Sql & "   CLIENTE.UF , " & Chr(10)
    Sql = Sql & "   CLIENTE.CEP , " & Chr(10)
    Sql = Sql & "   CLIENTE.NIVEL_CLIENTE , " & Chr(10)
    Sql = Sql & "   CLIENTE.INSCRICAO_ESTADUAL , " & Chr(10)
    Sql = Sql & "   CLIENTE.COBRANCA , " & Chr(10)
    Sql = Sql & "   CLIENTE.ENTREGA , " & Chr(10)
    Sql = Sql & "   CLIENTE.COMERCIAL,  " & Chr(10)
    Sql = Sql & "   CLIENTE.STATUS  " & Chr(10)
    
    Sql = Sql & "   FROM CLIENTE CLIENTE " & Chr(10)

    If Not Vazio(Trim(strUN_Negocio)) Then
         Sql = Sql & " INNER JOIN ATIVIDADES_CLIENTE ATIV_CLIENTE" & Chr(10)
         Sql = Sql & " ON(CLIENTE.CLIENTE = ATIV_CLIENTE.CLIENTE " & Chr(10)
         Sql = Sql & " AND ATIV_CLIENTE.UNIDADE_NEGOCIO = '" & Substitui(strUN_Negocio) & "')" & Chr(10)
    End If

    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CLIENTE.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCNPJ)) Then
         Sql = Sql & SqlAux & " CLIENTE.CNPJ = '" & Substitui(strCNPJ) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & " (CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strUF)) Then
         Sql = Sql & SqlAux & " CLIENTE.UF = '" & Substitui(strUF) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_CIDADE)) Then
         Sql = Sql & SqlAux & " CLIENTE.CIDADE LIKE '%" & Substitui(strLIKE_CIDADE) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
       
    If Not Vazio(Trim(strStatus)) Then
         Sql = Sql & SqlAux & " CLIENTE.STATUS = '" & Substitui(strStatus) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strNIVEL_CLIENTE)) Then
         Sql = Sql & SqlAux & " CLIENTE.NIVEL_CLIENTE = '" & Substitui(strNIVEL_CLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strAtendido)) Then
       If strAtendido <> "A" Then
         Sql = Sql & SqlAux & " CLIENTE.ATENDIDO = '" & Substitui(strAtendido) & "'" & Chr(10)
         SqlAux = " AND "
       End If
    End If
    
    Sql = Sql & "ORDER BY CLIENTE.RAZAO_SOCIAL " & Chr(10)
           
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
        
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_Cliente"), Err.Description

End Sub



Public Sub Relatorio_PedidoNaoLiquidado(ByVal strCOMPUTADOR As String, _
                                        ByVal strDATA_INICIAL As String, _
                                        ByVal strDATA_FINAL As String, _
                                        Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                        Optional ByVal strCOD_ARTIGO_PADRAO_LIKE As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_PEDIDO WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
        
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_PEDIDO"
       
    Sql = " INSERT INTO DB_IFV.dbo.TMP_PEDIDO " & Chr(10)
    Sql = Sql & "(COMPUTADOR, " & Chr(10)
    Sql = Sql & " NUM_PEDIDO," & Chr(10)
    Sql = Sql & " ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " CLIENTE," & Chr(10)
    Sql = Sql & " DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " POSTO," & Chr(10)
    Sql = Sql & " TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " COD_PRODUTO," & Chr(10)
    Sql = Sql & " QTD_PEDIDA," & Chr(10)
    Sql = Sql & " COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " DES_ARTIGO_PADRAO)" & Chr(10)
                
    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " PED.CLIENTE," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.POSTO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " ITENS.COD_PRODUTO," & Chr(10)
    Sql = Sql & " ITENS.QTD_PEDIDA," & Chr(10)
    Sql = Sql & " PROD.COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " ART.DES_ARTIGO_PADRAO" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    Sql = Sql & " INNER JOIN ITENS_PEDIDO ITENS ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN ARTIGO_PADRAO ART ON (ART.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO)" & Chr(10)
    Sql = Sql & " WHERE PED.COD_SIT_PEDIDO NOT IN ('F ', 'C ')" & Chr(10)

    SqlAux = " AND " & Chr(10)
    
    Sql = Sql & " AND (PED.DATA_REQUERIDA >= " & ToData(strDATA_INICIAL) & Chr(10)
    Sql = Sql & " AND PED.DATA_REQUERIDA <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
    
               
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " ART.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO_LIKE)) Then
         Sql = Sql & SqlAux & "( ART.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR  ART.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
               
               
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
        
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_PedidoNaoLiquidado"), Err.Description

End Sub

Public Sub Relatorio_PedidoCancelado(ByVal strCOMPUTADOR As String, _
                                     ByVal strDATA_INICIAL As String, _
                                     ByVal strDATA_FINAL As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_PEDIDO WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
        
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_PEDIDO"
        
    Sql = " INSERT INTO DB_IFV.dbo.TMP_PEDIDO " & Chr(10)
    Sql = Sql & "(COMPUTADOR, " & Chr(10)
    Sql = Sql & " NUM_PEDIDO," & Chr(10)
    Sql = Sql & " ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " CLIENTE," & Chr(10)
    Sql = Sql & " DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " POSTO," & Chr(10)
    Sql = Sql & " TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " RAZAO_SOCIAL)" & Chr(10)
                    
    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " PED.CLIENTE," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.POSTO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    Sql = Sql & " WHERE PED.COD_SIT_PEDIDO = 'C '" & Chr(10)

    SqlAux = " AND " & Chr(10)
    
    Sql = Sql & " AND (PED.DATA_REQUERIDA >= " & ToData(strDATA_INICIAL) & Chr(10)
    Sql = Sql & " AND PED.DATA_REQUERIDA <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
               
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
           
        
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_PedidoCancelado"), Err.Description

End Sub

Public Sub Relatorio_Produto(ByVal strCOMPUTADOR As String, _
                             Optional ByVal strIND_PROD As String, _
                             Optional ByVal strIND_COR As String, _
                             Optional ByVal strCOD_EMBALAGEM As String, _
                             Optional ByVal strCOD_EMBALAGEM_LIKE As String, _
                             Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                             Optional ByVal strCOD_ARTIGO_PADRAO_LIKE As String)
                              
    
    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_PRODUTO WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
        
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_PRODUTO"
        
    Sql = " INSERT INTO DB_IFV.dbo.TMP_PRODUTO " & Chr(10)
    Sql = Sql & "(COMPUTADOR, " & Chr(10)
    Sql = Sql & "COD_PRODUTO," & Chr(10)
    Sql = Sql & "DES_PRODUTO_CURTA," & Chr(10)
    Sql = Sql & "IND_ATIVO_PROD," & Chr(10)
    Sql = Sql & "COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & "DES_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & "COD_EMBALAGEM," & Chr(10)
    Sql = Sql & "DES_EMBALAGEM," & Chr(10)
    Sql = Sql & "COD_COR," & Chr(10)
    Sql = Sql & "IND_ATIVO_COR)" & Chr(10)
    
    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & "PROD.COD_PRODUTO," & Chr(10)
    Sql = Sql & "PROD.DES_PRODUTO_CURTA," & Chr(10)
    Sql = Sql & "PROD.IND_ATIVO AS IND_ATIVO_PROD," & Chr(10)
    Sql = Sql & "PROD.COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & "ART.DES_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & "PROD.COD_EMBALAGEM," & Chr(10)
    Sql = Sql & "EMB.DES_EMBALAGEM," & Chr(10)
    Sql = Sql & "PROD.COD_COR," & Chr(10)
    Sql = Sql & "COR.IND_ATIVO AS IND_ATIVO_COR" & Chr(10)
    Sql = Sql & "FROM PRODUTO PROD" & Chr(10)
    Sql = Sql & "LEFT JOIN COR ON (COR.COD_COR = PROD.COD_COR)" & Chr(10)
    Sql = Sql & "LEFT JOIN EMBALAGEM EMB ON (EMB.COD_EMBALAGEM = PROD.COD_EMBALAGEM)" & Chr(10)
    Sql = Sql & "LEFT JOIN ARTIGO_PADRAO ART ON (ART.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO)" & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strIND_PROD)) Then
         Sql = Sql & SqlAux & " PROD.IND_ATIVO = '" & Substitui(strIND_PROD) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strIND_COR)) Then
         Sql = Sql & SqlAux & " COR.IND_ATIVO = '" & Substitui(strIND_COR) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " ART.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO_LIKE)) Then
         Sql = Sql & SqlAux & "( ART.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR  ART.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strCOD_EMBALAGEM)) Then
         Sql = Sql & SqlAux & " EMB.COD_EMBALAGEM = '" & Substitui(strCOD_EMBALAGEM) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_EMBALAGEM_LIKE)) Then
         Sql = Sql & SqlAux & "( EMB.COD_EMBALAGEM LIKE '%" & Substitui(strCOD_EMBALAGEM_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR  EMB.DES_EMBALAGEM LIKE '%" & Substitui(strCOD_EMBALAGEM_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
               
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
           
        
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_Produto"), Err.Description

End Sub



Public Sub Relatorio_PedidoPendencia(ByVal strCOMPUTADOR As String, _
                                     ByVal strDATA_INICIAL As String, _
                                     ByVal strDATA_FINAL As String, _
                                     Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                     Optional ByVal strCOD_ARTIGO_PADRAO_LIKE As String, _
                                     Optional ByVal strCLIENTE As String, _
                                     Optional ByVal strCLIENTE_LIKE As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_PEDIDO WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
        
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_PEDIDO"
       
    Sql = " INSERT INTO DB_IFV.dbo.TMP_PEDIDO " & Chr(10)
    Sql = Sql & "(COMPUTADOR, " & Chr(10)
    Sql = Sql & " NUM_PEDIDO," & Chr(10)
    Sql = Sql & " ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " CLIENTE," & Chr(10)
    Sql = Sql & " DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " POSTO," & Chr(10)
    Sql = Sql & " TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " COD_PRODUTO," & Chr(10)
    Sql = Sql & " QTD_PEDIDA," & Chr(10)
    Sql = Sql & " VLR_PRECO," & Chr(10)
    Sql = Sql & " COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " DES_ARTIGO_PADRAO)" & Chr(10)
                
    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " PED.CLIENTE," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.POSTO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " ITENS.COD_PRODUTO," & Chr(10)
    Sql = Sql & " ITENS.QTD_PEDIDA," & Chr(10)
    Sql = Sql & " ITENS.VLR_PRECO," & Chr(10)
    Sql = Sql & " PROD.COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " ART.DES_ARTIGO_PADRAO" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    Sql = Sql & " INNER JOIN ITENS_PEDIDO ITENS ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN ARTIGO_PADRAO ART ON (ART.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO)" & Chr(10)
    Sql = Sql & " WHERE PED.NUM_PEDIDO IN (SELECT NUM_PEDIDO FROM PEDIDO_BLOQUEIOS WHERE STATUS = 'A')" & Chr(10)

    SqlAux = " AND " & Chr(10)
    
    Sql = Sql & " AND (PED.DATA_REQUERIDA >= " & ToData(strDATA_INICIAL) & Chr(10)
    Sql = Sql & " AND PED.DATA_REQUERIDA <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
    
               
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " ART.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO_LIKE)) Then
         Sql = Sql & SqlAux & "( ART.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR  ART.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CLI.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
            
    If Not Vazio(Trim(strCLIENTE_LIKE)) Then
         Sql = Sql & SqlAux & " (CLI.CNPJ LIKE '%" & Substitui(strCLIENTE_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR CLI.RAZAO_SOCIAL LIKE '%" & Substitui(strCLIENTE_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
                              
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
        
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_PedidoPendencia"), Err.Description

End Sub



Public Sub Relatorio_Vendas_Diarias(ByVal strCOMPUTADOR As String, _
                                    ByVal strDATA_INICIAL As String, _
                                    ByVal strDATA_FINAL As String, _
                                    Optional ByVal strCOD_ARTIGO_PADRAO As String, _
                                    Optional ByVal strCOD_ARTIGO_PADRAO_LIKE As String, _
                                    Optional ByVal strCLIENTE As String, _
                                    Optional ByVal strCLIENTE_LIKE As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_PEDIDO WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
    
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_PEDIDO"
        
    Sql = " INSERT INTO DB_IFV.dbo.TMP_PEDIDO " & Chr(10)
    Sql = Sql & "(COMPUTADOR, " & Chr(10)
    Sql = Sql & " NUM_PEDIDO," & Chr(10)
    Sql = Sql & " ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " CLIENTE," & Chr(10)
    Sql = Sql & " DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " POSTO," & Chr(10)
    Sql = Sql & " TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " DATA_REQUERIDA," & Chr(10)
    Sql = Sql & " RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " COD_PRODUTO," & Chr(10)
    Sql = Sql & " QTD_PEDIDA," & Chr(10)
    Sql = Sql & " COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " DES_ARTIGO_PADRAO)" & Chr(10)
                
    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & " PED.NUM_PEDIDO," & Chr(10)
    Sql = Sql & " PED.ESTABELECIMENTO," & Chr(10)
    Sql = Sql & " PED.CLIENTE," & Chr(10)
    Sql = Sql & " PED.DATA_INCLUSAO," & Chr(10)
    Sql = Sql & " PED.DES_TRANSPORTE," & Chr(10)
    Sql = Sql & " PED.DATA_ULT_ALT," & Chr(10)
    Sql = Sql & " PED.PEDIDO_CLIENTE," & Chr(10)
    Sql = Sql & " PED.POSTO," & Chr(10)
    Sql = Sql & " PED.TOTAL_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_COND_COMERCIAL," & Chr(10)
    Sql = Sql & " PED.DES_SIT_FATUR," & Chr(10)
    Sql = Sql & " PED.COD_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " PED.DES_SIT_PEDIDO," & Chr(10)
    Sql = Sql & " ITENS.DATA_CEDO," & Chr(10)
    Sql = Sql & " CLI.RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & " ITENS.COD_PRODUTO," & Chr(10)
    Sql = Sql & " ITENS.QTD_PEDIDA," & Chr(10)
    Sql = Sql & " PROD.COD_ARTIGO_PADRAO," & Chr(10)
    Sql = Sql & " ART.DES_ARTIGO_PADRAO" & Chr(10)
    Sql = Sql & " FROM PEDIDO PED" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN CLIENTE CLI ON (CLI.CLIENTE= PED.CLIENTE)" & Chr(10)
    Sql = Sql & " INNER JOIN ITENS_PEDIDO ITENS ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN PRODUTO PROD ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO)" & Chr(10)
    Sql = Sql & " LEFT OUTER JOIN ARTIGO_PADRAO ART ON (ART.COD_ARTIGO_PADRAO = PROD.COD_ARTIGO_PADRAO)" & Chr(10)
    ' Sql = Sql & " WHERE PED.COD_SIT_PEDIDO = 'F '" & Chr(10)

    SqlAux = " AND " & Chr(10)
    
  '  Sql = Sql & " WHERE (PED.DATA_REQUERIDA >= " & ToData(strDATA_INICIAL) & Chr(10)
  '  Sql = Sql & " AND PED.DATA_REQUERIDA <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
  '  Sql = Sql & " AND PED.COD_SIT_PEDIDO <> 'C '" & Chr(10)
    
    Sql = Sql & " WHERE (ITENS.DATA_CEDO >= " & ToData(strDATA_INICIAL) & Chr(10)
    Sql = Sql & " AND ITENS.DATA_TARDE <= " & ToData(strDATA_FINAL) & ")" & Chr(10)
    Sql = Sql & " AND ITENS.SITUACAO <> 'C '" & Chr(10)
    Sql = Sql & " AND PED.COD_SIT_PEDIDO <> 'C '" & Chr(10)
               
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO)) Then
         Sql = Sql & SqlAux & " ART.COD_ARTIGO_PADRAO = '" & Substitui(strCOD_ARTIGO_PADRAO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strCOD_ARTIGO_PADRAO_LIKE)) Then
         Sql = Sql & SqlAux & "( ART.COD_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR  ART.DES_ARTIGO_PADRAO LIKE '%" & Substitui(strCOD_ARTIGO_PADRAO_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    
    If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " CLI.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
            
    If Not Vazio(Trim(strCLIENTE_LIKE)) Then
         Sql = Sql & SqlAux & " (CLI.CNPJ LIKE '%" & Substitui(strCLIENTE_LIKE) & "%'" & Chr(10)
         Sql = Sql & " OR CLI.RAZAO_SOCIAL LIKE '%" & Substitui(strCLIENTE_LIKE) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
               
               
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
           
        
    Exit Sub

ErrorHandler:
    
    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_Vendas_Diarias"), Err.Description

End Sub

Public Sub Relatorio_Duplicatas(ByVal strCOMPUTADOR As String, _
                                Optional ByVal strCOD_NOTA_FISCAL As String, _
                                Optional ByVal strSERIE As String, _
                                Optional ByVal strSEQUENCIA As String, _
                                Optional ByVal strNUM_PEDIDO As String, _
                                Optional ByVal strCLIENTE As String, _
                                Optional ByVal strLIKE_CNPJ_NOME_FANTASIA As String, _
                                Optional ByVal strDATA_VENCIMENTO_INICIAL As String, _
                                Optional ByVal strDATA_VENCIMENTO_FINAL As String, _
                                Optional ByVal strDATA_PAGAMENTO_INICIAL As String, _
                                Optional ByVal strDATA_PAGAMENTO_FINAL As String, _
                                Optional ByVal strStatus As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String
    
    
    If Not Vazio(Trim(strDATA_VENCIMENTO_INICIAL)) And Not Vazio(Trim(strDATA_VENCIMENTO_FINAL)) Then
    
        If DateValue(strDATA_VENCIMENTO_INICIAL) > DateValue(strDATA_VENCIMENTO_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Relatorio_Duplicatas"), "Data Vencimento Final deve conter valor maior ou igual Data Vencimento In cio."
        End If
        
    End If
    
    If Not Vazio(Trim(strDATA_PAGAMENTO_INICIAL)) And Not Vazio(Trim(strDATA_PAGAMENTO_FINAL)) Then
    
        If DateValue(strDATA_PAGAMENTO_INICIAL) > DateValue(strDATA_PAGAMENTO_FINAL) Then
            Err.Raise vbObjectError, SetErrSource(strNomeModulo, "Relatorio_Duplicatas"), "Data Pagamento Final deve conter valor maior ou igual Data Pagamento In cio."
        End If
        
    End If

    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_DUPLICATAS WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
    
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_DUPLICATAS"
   
    Sql = "INSERT INTO DB_IFV.dbo.TMP_DUPLICATAS (" & Chr(10)
    Sql = Sql & "COMPUTADOR," & Chr(10)
    Sql = Sql & "COD_NOTA_FISCAL," & Chr(10)
    Sql = Sql & "SERIE," & Chr(10)
    Sql = Sql & "SEQUENCIA," & Chr(10)
    Sql = Sql & "DATA_VENCIMENTO," & Chr(10)
    Sql = Sql & "VALOR," & Chr(10)
    Sql = Sql & "DES_PORTADOR," & Chr(10)
    Sql = Sql & "VALOR_PAGAMENTO," & Chr(10)
    Sql = Sql & "VALOR_SALDO," & Chr(10)
    Sql = Sql & "DATA_PAGAMENTO," & Chr(10)
    Sql = Sql & "NUM_PEDIDO," & Chr(10)
    Sql = Sql & "CNPJ," & Chr(10)
    Sql = Sql & "RAZAO_SOCIAL," & Chr(10)
    Sql = Sql & "NOME_FANTASIA )" & Chr(10)

    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & "   DUPL.COD_NOTA_FISCAL , " & Chr(10)
    Sql = Sql & "   DUPL.SERIE , " & Chr(10)
    Sql = Sql & "   DUPL.SEQUENCIA , " & Chr(10)
    Sql = Sql & "   DUPL.DATA_VENCIMENTO, " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR,'0') VALOR, " & Chr(10)
    Sql = Sql & "   DUPL.DES_PORTADOR , " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR_PAGAMENTO,'0') VALOR_PAGAMENTO, " & Chr(10)
    Sql = Sql & "   ISNULL(DUPL.VALOR_SALDO,'0') VALOR_SALDO, " & Chr(10)
    Sql = Sql & "   DUPL.DATA_PAGAMENTO, " & Chr(10)
    Sql = Sql & "   NOT_FIS.NUM_PEDIDO, " & Chr(10)
    Sql = Sql & "   CLIENTE.CNPJ, " & Chr(10)
    Sql = Sql & "   CLIENTE.RAZAO_SOCIAL, " & Chr(10)
    Sql = Sql & "   CLIENTE.NOME_FANTASIA " & Chr(10)

    Sql = Sql & "   FROM DUPLICATAS DUPL " & Chr(10)
    Sql = Sql & "   INNER JOIN NOTA_FISCAL NOT_FIS" & Chr(10)
    Sql = Sql & "       ON(NOT_FIS.COD_NOTA_FISCAL = DUPL.COD_NOTA_FISCAL " & Chr(10)
    Sql = Sql & "          AND NOT_FIS.SERIE = DUPL.SERIE )" & Chr(10)
    Sql = Sql & "   LEFT OUTER JOIN CLIENTE CLIENTE " & Chr(10)
    Sql = Sql & "        ON (CLIENTE.CLIENTE = NOT_FIS.CLIENTE) " & Chr(10)
    
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strCOD_NOTA_FISCAL)) Then
         Sql = Sql & SqlAux & " DUPL.COD_NOTA_FISCAL = '" & Substitui(strCOD_NOTA_FISCAL) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSERIE)) Then
         Sql = Sql & SqlAux & " DUPL.SERIE = '" & Substitui(strSERIE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strSEQUENCIA)) Then
         Sql = Sql & SqlAux & " DUPL.SEQUENCIA = '" & Substitui(strSEQUENCIA) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strNUM_PEDIDO)) Then
         Sql = Sql & SqlAux & " NOT_FIS.NUM_PEDIDO = '" & Substitui(strNUM_PEDIDO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
     If Not Vazio(Trim(strCLIENTE)) Then
         Sql = Sql & SqlAux & " NOT_FIS.CLIENTE = '" & Substitui(strCLIENTE) & "'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strLIKE_CNPJ_NOME_FANTASIA)) Then
         Sql = Sql & SqlAux & "(CLIENTE.CNPJ LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%'" & Chr(10)
         Sql = Sql & " OR CLIENTE.RAZAO_SOCIAL LIKE '%" & Substitui(strLIKE_CNPJ_NOME_FANTASIA) & "%')" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_VENCIMENTO_INICIAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_VENCIMENTO >= " & ToData(strDATA_VENCIMENTO_INICIAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_VENCIMENTO_FINAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_VENCIMENTO <= " & ToData(strDATA_VENCIMENTO_FINAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_PAGAMENTO_INICIAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO >= " & ToData(strDATA_PAGAMENTO_INICIAL) & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strDATA_PAGAMENTO_FINAL)) Then
         Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO <= " & ToData(strDATA_PAGAMENTO_FINAL) & Chr(10)
         SqlAux = " AND "
    End If
   
    
    Select Case strStatus
            '  VENCER
            Case "A"
                Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO IS NULL " & Chr(10)
                Sql = Sql & " AND DUPL.DATA_VENCIMENTO >= " & ToData(Date) & Chr(10)
                SqlAux = " AND "
            'VENCIDO
            Case "V"
                Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO IS NULL " & Chr(10)
                Sql = Sql & " AND DUPL.DATA_VENCIMENTO < " & ToData(Date) & Chr(10)
                SqlAux = " AND "
            'LIQUIDADO
            Case "L"
                Sql = Sql & SqlAux & " DUPL.DATA_PAGAMENTO IS NOT NULL " & Chr(10)
                SqlAux = " AND "
    
    End Select
        
    Sql = Sql & "ORDER BY DUPL.COD_NOTA_FISCAL, DUPL.SERIE, DUPL.SEQUENCIA  " & Chr(10)
           
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
    
    Exit Sub

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_Duplicatas"), Err.Description

End Sub

Public Sub Relatorio_PrevCotaVen(ByVal strCOMPUTADOR As String, _
                                Optional ByVal strLIKE_FAMILIA As String, _
                                Optional ByVal strLIKE_MESANO As String, _
                                Optional ByVal strLIKE_LINHA_PROD As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String
    
    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_PREVISAO WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
    
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_PREVISAO"
    
    Sql = "INSERT INTO DB_IFV.dbo.TMP_PREVISAO (" & Chr(10)
    Sql = Sql & "COMPUTADOR," & Chr(10)
    Sql = Sql & "MESANO," & Chr(10)
    Sql = Sql & "LINHA_PRODUTO," & Chr(10)
    Sql = Sql & "FAMILIA," & Chr(10)
    Sql = Sql & "CATEGORIA," & Chr(10)
    Sql = Sql & "FABRICA," & Chr(10)
    Sql = Sql & "PREVISAO," & Chr(10)
    Sql = Sql & "COTA," & Chr(10)
    Sql = Sql & "VENDA )" & Chr(10)

    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & "   PREV_CT_V.MESANO , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.LINHA_PRODUTO , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.FAMILIA , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.CATEGORIA , " & Chr(10)
    Sql = Sql & "   PREV_CT_V.FABRICA , " & Chr(10)
    Sql = Sql & "   ISNULL(PREV_CT_V.PREVISAO,'0') PREVISAO , " & Chr(10)
    Sql = Sql & "   ISNULL(PREV_CT_V.COTA,'0')  COTA, " & Chr(10)
    Sql = Sql & "   ISNULL(PREV_CT_V.VENDA,'0')  VENDA " & Chr(10)
                
    Sql = Sql & "   FROM PREVISAO_COTA_VENDA PREV_CT_V " & Chr(10)
       
    SqlAux = " WHERE " & Chr(10)
    
    If Not Vazio(Trim(strLIKE_FAMILIA)) Then
         Sql = Sql & SqlAux & " PREV_CT_V.FAMILIA  LIKE '%" & Substitui(strLIKE_FAMILIA) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
    
    If Not Vazio(Trim(strLIKE_MESANO)) Then
         Sql = Sql & SqlAux & " PREV_CT_V.MESANO  LIKE '%" & Substitui(strLIKE_MESANO) & "%'" & Chr(10)
         SqlAux = " AND "
    End If
        
    If Not Vazio(Trim(strLIKE_LINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PREV_CT_V.LINHA_PRODUTO LIKE '" & Substitui(strLIKE_LINHA_PRODUTO) & "'" & Chr(10)
         SqlAux = " AND "
    End If
    
    Sql = Sql & "ORDER BY PREV_CT_V.MESANO, PREV_CT_V.LINHA_PRODUTO, PREV_CT_V.FAMILIA  " & Chr(10)
           
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
    
    Exit Sub

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_PrevCotaVen"), Err.Description

End Sub

Public Sub Relatorio_EstFamCota(ByVal strCOMPUTADOR As String, _
                                Optional ByVal strDTIMPR As String, _
                                Optional ByVal strLIKE_LINHA_PROD As String, _
                                Optional ByVal strQUALIDADE As String)

    On Error GoTo ErrorHandler

    Dim Sql            As String
    Dim Sql_Delete     As String
    Dim strConnect     As String
    Dim cnn            As adodb.Connection
    Dim SqlAux         As String
    
    strConnect = ConnectionSQL("IFV")
    
'    Sql_Delete = " DELETE DB_IFV.dbo.TMP_ESTFAMCOTA WHERE COMPUTADOR = '" & strCOMPUTADOR & "'"
    
    Sql_Delete = " TRUNCATE TABLE DB_IFV.dbo.TMP_ESTFAMCOTA"
    
    Sql = "INSERT INTO DB_IFV.dbo.TMP_ESTFAMCOTA (" & Chr(10)
    Sql = Sql & "COMPUTADOR," & Chr(10)
    Sql = Sql & "LINHA_PRODUTO," & Chr(10)
    Sql = Sql & "FAMILIA," & Chr(10)
    Sql = Sql & "MES," & Chr(10)
    Sql = Sql & "ANO," & Chr(10)
    Sql = Sql & "VOLUME )" & Chr(10)

    Sql = Sql & " SELECT " & Chr(10)
    Sql = Sql & "'" & strCOMPUTADOR & "', " & Chr(10)
    Sql = Sql & "   PROD.LINHA_PRODUTO, " & Chr(10)
    Sql = Sql & "   PROD.FAMILIA_COTA, " & Chr(10)
    Sql = Sql & "   RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2) MES_REQ, " & Chr(10)
    Sql = Sql & "   RIGHT('0' + LTRIM(YEAR(ITENS.DATA_CEDO)),4) ANO_REQ, " & Chr(10)
    Sql = Sql & "   ISNULL(SUM(ITENS.QTD_PEDIDA), 0 ) TOTAL_PEDIDO " & Chr(10)
    
    Sql = Sql & "   FROM PEDIDO PED " & Chr(10)
   
    Sql = Sql & "   LEFT OUTER JOIN ITENS_PEDIDO ITENS " & Chr(10)
    Sql = Sql & "        ON (ITENS.NUM_PEDIDO = PED.NUM_PEDIDO) " & Chr(10)
   
    Sql = Sql & "   LEFT OUTER JOIN PRODUTO PROD " & Chr(10)
    Sql = Sql & "        ON (PROD.COD_PRODUTO = ITENS.COD_PRODUTO) " & Chr(10)
   
    SqlAux = " WHERE " & Chr(10)

    Sql = Sql & SqlAux & " PED.COD_SIT_PEDIDO <> 'C' " & Chr(10)
    
    SqlAux = " AND "
    
    If Not Vazio(Trim(strDTIMPR)) Then
       Sql = Sql & SqlAux & " MONTH(ITENS.DATA_CEDO) = MONTH('" & Substitui(strDTIMPR) & "')" & Chr(10)
       Sql = Sql & SqlAux & " YEAR(ITENS.DATA_CEDO) = YEAR('" & Substitui(strDTIMPR) & "')" & Chr(10)
    End If
    
    Sql = Sql & SqlAux & " ITENS.SITUACAO <> 'C' " & Chr(10)
    
    Sql = Sql & SqlAux & " ITENS.QTD_PEDIDA Is Not Null " & Chr(10)
    
    If Not Vazio(Trim(strQUALIDPROD)) Then
         Sql = Sql & SqlAux & " PROD.COD_QUALIDADE = '" & Substitui(strQUALIDPROD) & "'" & Chr(10)
    End If
    
    If Not Vazio(Trim(strLINHA_PRODUTO)) Then
         Sql = Sql & SqlAux & " PROD.LINHA_PRODUTO LIKE '" & Substitui(strLINHA_PRODUTO) & "'" & Chr(10)
    End If
        
    Sql = Sql & "GROUP BY PROD.LINHA_PRODUTO, PROD.FAMILIA_COTA, RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2), RIGHT('0' + LTRIM(YEAR(ITENS.DATA_CEDO)),4) " & Chr(10)
    
    Sql = Sql & "ORDER BY PROD.LINHA_PRODUTO, PROD.FAMILIA_COTA, RIGHT('0' + LTRIM(MONTH(ITENS.DATA_CEDO)),2), RIGHT('0' + LTRIM(YEAR(ITENS.DATA_CEDO)),4) " & Chr(10)
           
    Set cnn = New adodb.Connection

    With cnn
        .Open strConnect
        .Execute Sql_Delete, , adExecuteNoRecords
        .Execute Sql, , adExecuteNoRecords
    End With
    
    Exit Sub

ErrorHandler:

    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

    Err.Raise Err.Number, SetErrSource(strNomeModulo, "Relatorio_EstFamCota"), Err.Description

End Sub




## Modulo_Generico.bas
Attribute VB_Name = "Modulo_Generico"
Option Explicit

'--------------------------------'
'- Definicao de Tipos Genericos -'
'--------------------------------'

'Tipo utilizado para Informar o Status de uma Janela de Manutencao
Public Enum conStatus_Manut
   conConsultando = 0
   conIncluindo = 1
   conEditando = 2
End Enum


'Variavel do tipo Form Generico para ser utilizado por toda a aplicacao
Public gfrmForm As Form
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'-------------------------------'
'- Rotinas e Funcoes Genericas -'
'-------------------------------'

Public Sub Centra_Form(frmFORM As Form,  MDI As Boolean)
   If  MDI Then
'      frmFORM.Left = (frmMain.ScaleWidth / 2) - (frmFORM.ScaleWidth / 2)
'      frmFORM.Top = (frmMain.ScaleHeight / 2) - (frmFORM.ScaleHeight / 2)
   Else
      frmFORM.Left = (Screen.Width / 2) - (frmFORM.ScaleWidth / 2)
      frmFORM.Top = (Screen.Height / 2) - (frmFORM.ScaleHeight / 2)
   End If
End Sub


Public Function Informa_String_Status(pStatus_Manut As conStatus_Manut) As String
   Select Case pStatus_Manut
      Case conConsultando
         Informa_String_Status = " - [Consultando]"
      Case conIncluindo
         Informa_String_Status = " - [Incluindo]"
      Case conEditando
         Informa_String_Status = " - [Editando/Alterando]"
   End Select
End Function

Function ComputerName() As String
    Dim mlngTemp As Long
    Dim mstrComp As String
    mstrComp = String(145, Chr(0))
    mlngTemp = GetComputerName(mstrComp, 145)
    ComputerName = Left(mstrComp, InStr(mstrComp, Chr(0)) - 1)
End Function






