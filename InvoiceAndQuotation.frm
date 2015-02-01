VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form InvoiceAndQuotation 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10050
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "InvoiceAndQuotation.frx":0000
   ScaleHeight     =   10050
   ScaleWidth      =   15210
   Begin VB.PictureBox FrmOperative 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      Picture         =   "InvoiceAndQuotation.frx":26EAE
      ScaleHeight     =   3495
      ScaleWidth      =   14775
      TabIndex        =   39
      Top             =   5880
      Width           =   14775
      Begin VB.PictureBox FrmDate 
         Height          =   2895
         Left            =   8040
         Picture         =   "InvoiceAndQuotation.frx":4DD5C
         ScaleHeight     =   2835
         ScaleWidth      =   6555
         TabIndex        =   55
         Top             =   600
         Width           =   6615
         Begin VB.PictureBox frmPagos 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   0
            Picture         =   "InvoiceAndQuotation.frx":74C0A
            ScaleHeight     =   1815
            ScaleWidth      =   3735
            TabIndex        =   59
            Top             =   720
            Width           =   3735
            Begin VB.TextBox TxtCast 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   405
               Left            =   1680
               TabIndex        =   60
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label LblValue_CastSugery 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Microsoft JhengHei UI"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   1680
               TabIndex        =   65
               Top             =   120
               Width           =   2055
            End
            Begin VB.Label LblCastSugery 
               BackStyle       =   0  'Transparent
               Caption         =   "Abono sugerido 50%"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   0
               TabIndex        =   64
               Top             =   0
               Width           =   1695
            End
            Begin VB.Label LblCast 
               BackStyle       =   0  'Transparent
               Caption         =   "Digite valor del abono"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   0
               TabIndex        =   63
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label LblValue_Sald 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Microsoft JhengHei UI"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   1680
               TabIndex        =   62
               Top             =   1320
               Width           =   2055
            End
            Begin VB.Label LblSald 
               BackStyle       =   0  'Transparent
               Caption         =   "SALDO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   0
               TabIndex        =   61
               Top             =   1320
               Width           =   1695
            End
         End
         Begin VB.TextBox TxtDaysEnd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   405
            Left            =   2760
            TabIndex        =   56
            Top             =   120
            Width           =   975
         End
         Begin MSComCtl2.MonthView DPSale 
            Height          =   2370
            Left            =   3840
            TabIndex        =   57
            Top             =   120
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   4194304
            BackColor       =   4194304
            Appearance      =   1
            MonthBackColor  =   16777215
            StartOfWeek     =   52756481
            TitleBackColor  =   4194304
            TitleForeColor  =   16777215
            TrailingForeColor=   4194304
            CurrentDate     =   42022
         End
         Begin VB.Label LblDateEnd 
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo de Fabricacion (Dias)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   2295
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridList_Operative 
         Height          =   1215
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         ForeColor       =   4194304
         BackColorFixed  =   4194304
         ForeColorFixed  =   16777215
         BackColorSel    =   16776960
         ForeColorSel    =   4194304
         BackColorBkg    =   4194304
         GridColor       =   4194304
         GridColorFixed  =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblSubTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   54
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label LblValue_Subtotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   5880
         TabIndex        =   53
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label LblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   52
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label LblValue_Total 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   5880
         TabIndex        =   51
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label LblIva 
         BackStyle       =   0  'Transparent
         Caption         =   "IVA %"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   50
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label LblValue_Iva 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   5880
         TabIndex        =   49
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label LblValue_Date 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12600
         TabIndex        =   48
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label LblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11400
         TabIndex        =   47
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label LblValue_Neto 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2160
         TabIndex        =   46
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label LblV_Neto 
         BackStyle       =   0  'Transparent
         Caption         =   "V. materiales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   45
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label LdValue_Double 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2160
         TabIndex        =   44
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label LdlDouble 
         BackStyle       =   0  'Transparent
         Caption         =   "Mano de Obra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   43
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label LblWinner 
         BackStyle       =   0  'Transparent
         Caption         =   "Ganancia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   42
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label LblValue_Winner 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   2640
         Width           =   2055
      End
   End
   Begin VB.PictureBox FrmCapture 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      Picture         =   "InvoiceAndQuotation.frx":9BAB8
      ScaleHeight     =   2175
      ScaleWidth      =   14775
      TabIndex        =   26
      Top             =   3600
      Width           =   14775
      Begin VB.ComboBox CbnImputs 
         Height          =   315
         Left            =   1080
         TabIndex        =   31
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox CbnMeasure 
         Height          =   315
         Left            =   5640
         TabIndex        =   30
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   4455
      End
      Begin VB.OptionButton OpYes 
         BackColor       =   &H00400000&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12120
         Picture         =   "InvoiceAndQuotation.frx":C2966
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton OpNot 
         BackColor       =   &H00400000&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   13200
         Picture         =   "InvoiceAndQuotation.frx":C4059
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtQuanty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   525
         Left            =   11640
         TabIndex        =   27
         Top             =   120
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid GridList_Input 
         Height          =   1215
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         ForeColor       =   4194304
         BackColorFixed  =   4194304
         ForeColorFixed  =   16777215
         BackColorSel    =   16776960
         ForeColorSel    =   4194304
         BackColorBkg    =   4194304
         GridColor       =   4194304
         GridColorFixed  =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblMaterials 
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LblMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "Medida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblRequest 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¿Requiere IVA?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11760
         TabIndex        =   34
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label LblQuanty 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10560
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox FrmBody 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      Picture         =   "InvoiceAndQuotation.frx":C574C
      ScaleHeight     =   1095
      ScaleWidth      =   14775
      TabIndex        =   13
      Top             =   960
      Width           =   14775
      Begin VB.TextBox TxtObservations 
         Height          =   375
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   8175
      End
      Begin VB.TextBox TxtTypeDocument 
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox TxtEmail 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TxtAddress 
         Height          =   375
         Left            =   10920
         TabIndex        =   16
         Top             =   120
         Width           =   3735
      End
      Begin VB.TextBox TxtPhone 
         Height          =   375
         Left            =   7800
         TabIndex        =   15
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox TxtDocument 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label LblObservations 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label LblDocumentNumber 
         BackStyle       =   0  'Transparent
         Caption         =   " Documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electronico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label LblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10080
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblPhone 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox FrmDescription 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      Picture         =   "InvoiceAndQuotation.frx":EC5FA
      ScaleHeight     =   855
      ScaleWidth      =   14775
      TabIndex        =   9
      Top             =   2400
      Width           =   14775
      Begin VB.TextBox TxtDescripProject 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   9015
      End
      Begin VB.Label LblNumber 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12480
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Lbltitle_in 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9720
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.PictureBox FrmClient 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   360
      Picture         =   "InvoiceAndQuotation.frx":1134A8
      ScaleHeight     =   735
      ScaleWidth      =   11175
      TabIndex        =   4
      Top             =   -120
      Width           =   11175
      Begin VB.CommandButton BtnSearch 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "InvoiceAndQuotation.frx":13A356
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CbnSearch 
         Height          =   315
         ItemData        =   "InvoiceAndQuotation.frx":13BA49
         Left            =   2640
         List            =   "InvoiceAndQuotation.frx":13BA4B
         TabIndex        =   7
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   5865
      End
      Begin VB.OptionButton OpName 
         BackColor       =   &H00400000&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   1320
         Picture         =   "InvoiceAndQuotation.frx":13BA4D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Opdoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Picture         =   "InvoiceAndQuotation.frx":13D140
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "InvoiceAndQuotation.frx":13E833
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9480
      Width           =   2295
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      Picture         =   "InvoiceAndQuotation.frx":13FF26
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton BtnCreateClient 
      Caption         =   "CREAR CLIENTE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      Picture         =   "InvoiceAndQuotation.frx":141619
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Lbltittledescrip 
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCIÓN DEL PROYECTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   2040
      Width           =   10215
   End
   Begin VB.Label LblTittleCapture 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   3240
      Width           =   10215
   End
   Begin VB.Label LbltittleInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   10215
   End
   Begin VB.Label LblhelpGeneral 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   9480
      Width           =   11655
   End
End
Attribute VB_Name = "InvoiceAndQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''' REGION GLOBALES
Public Crear_client As Integer
Public Count_GOperative As Long
Public GValueNeto As Long
Public GValueDouble As Long
Public GValueWinner As Long
Public GValueSubTotal As Long
Public GValueIva As Long
Public GValueTotal As Long
Public GVAbono As Long
Public GSwitch_Delete As Integer
Public GType_operation As String
Public GName As String
'''''' END_REGION GLOBALES

''''''----------- REGION ENVENTOS
'INICIO DE FORM operativo
Private Sub Form_Load()

    Dim Inputs() As Variant
    Dim C_Proc As New C_General_Procedures
    
    Crear_client = 0
    FrmBody.Visible = False
    FrmCapture.Visible = False
    FrmDescription.Visible = False
    FrmOperative.Visible = False
    LblValue_Date.Caption = Date
    LblhelpGeneral.Visible = False
    LblTittleCapture.Visible = False
    Lbltittledescrip.Visible = False
    BtnCreate.Visible = False
    CbnSearch.BackColor = &HFFFFFF
    
    CbnMeasure.Enabled = False
    Count_GOperative = 1
    
    'cargamos consulta datos insumos en BD
    Inputs = C_Proc.Datos_Charge("TC_Inputs", "Charge")
    
    'traemos la cantidad de insumos en BD
    Q_Inputs = C_Proc.Q_Combo("TC_Inputs")
    Q_Inputs = Q_Inputs - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_Inputs
           CbnImputs.AddItem Inputs(1, I)
    Next
      
    
    'dimencionamos el numero de columnas del grid
    GridList_Input.Cols = GridList_Input.Cols + 4
    
    'cargo titulos del grid
    GridList_Input.TextMatrix(0, 0) = "Proveedor"
    GridList_Input.TextMatrix(0, 1) = "Material"
    GridList_Input.TextMatrix(0, 2) = "Medida"
    GridList_Input.TextMatrix(0, 3) = "Descripción"
    GridList_Input.TextMatrix(0, 4) = "Valor"
    GridList_Input.TextMatrix(0, 5) = "Seleccionar"
    
    GridList_Input.Visible = False
    
    GridList_Operative.Cols = GridList_Operative.Cols + 3
    
    GridList_Operative.TextMatrix(0, 0) = "Descripción"
    GridList_Operative.TextMatrix(0, 1) = "Valor Unidad"
    GridList_Operative.TextMatrix(0, 2) = "Cantidad"
    GridList_Operative.TextMatrix(0, 3) = "Valor Total"
    GridList_Operative.TextMatrix(0, 4) = "Eliminar"
    
End Sub
'FIN DE FORM operativo

'boton salir
Private Sub BtnExit_Click()
     Unload InvoiceAndQuotation
End Sub

'boton crear cliente
Private Sub BtnCreateClient_Click()
    
    Crear_client = 1
    Load Client_Crud
    Client_Crud.Left = (MenuCarpenter.ScaleWidth - Client_Crud.Width) / 2
    Client_Crud.Top = (MenuCarpenter.ScaleHeight - Client_Crud.Height) / 2
    Client_Crud.Caption = "Nuevo Cliente"
    Client_Crud.FrmClient.Visible = False
    Client_Crud.BtnCreate.Caption = "CREAR CLIENTE"
    Client_Crud.Show

End Sub

'boton buscar cliente
Private Sub BtnSearch_Click()

    Dim C_Client As New C_CRUD_client
    Dim C_Project As New C_Project
    Dim Traer_Datos() As Variant
    Dim op_Search As String
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
      
    FrmBody.Visible = True
          
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(1)
      
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar la " & GType_operation & " Debe almenos seleccionar un cliente por favor!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        ' traer los datos del registro seleccionado
        Traer_Datos = C_Client.Charge_List_Client(op_Search, CbnSearch.Text)
            
        'cargar los datos capturados en los campos
        
        TxtDocument.Text = Traer_Datos(2, 0)
        TxtTypeDocument.Text = Traer_Datos(1, 0)
        TxtAddress.Text = Traer_Datos(3, 0)
        TxtPhone.Text = Traer_Datos(4, 0)
        TxtEmail.Text = Traer_Datos(5, 0)
        TxtObservations.Text = Traer_Datos(6, 0)
        G_IdDocument = Traer_Datos(7, 0)
        LbltittleInfo.Caption = "INFORMACIÓN PRINCIPAL DE " & Traer_Datos(0, 0)
        GName = Traer_Datos(0, 0)
        FrmBody.Visible = True
        FrmCapture.Visible = True
        FrmDescription.Visible = True
        LblTittleCapture.Visible = True
        Lbltittledescrip.Visible = True
        block
        
        'capturamos si hay factura o no
        Q_Fact = C_Project.Q_Project(GType_operation)
        
        If Q_Fact = 0 Then
         LblNumber.Caption = 1
        Else
          Q_Fact = Q_Fact + 1
          LblNumber.Caption = Q_Fact
        End If
        
    End If

End Sub

'boton genanerar factura o cotizacion
Private Sub BtnCreate_Click()

    Dim validate As Integer
    
    'validamos campos de diligenciamiento
    validate = ValidateCampos(2)
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar la " & GType_operation & " Debe llenar los campos en Rojo!!!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
    Select Case GType_operation
    
        Case "Factura"
            G_Factura
        
        Case "Cotización"
            G_Cotizacion
            
        Case Else
        
    End Select
    
      
    End If

End Sub

'para volver a cargar el combo de clientes si se creo uno nuevo
Private Sub CbnSearch_GotFocus()
    
    If Crear_client = 1 Then
         If OpName.Value = True Then
            Call OpName_Click
        Else
            Call Opdoc_Click
        End If
    End If
    
End Sub

'para desbloquear el combo de medidas
Private Sub CbnImputs_GotFocus()
    CbnMeasure.Enabled = True
End Sub

'para cargar el combo de medidas segun el insumo
Private Sub CbnImputs_lostFocus()

   Dim C_ListInt As New C_List_Price
   Dim IndexInputs As Integer
   Dim ListInt() As Variant
    
   CbnMeasure.clear
   
   IndexInputs = CbnImputs.ListIndex
   
   'cargamos consulta datos medidas según el insumo solicitado en BD
   ListInt = C_ListInt.Measure(IndexInputs)
   
   'traemos la cantidad de medidas según el insumo solicitado en BD
   Q_Measure = C_ListInt.Q_Measure(IndexInputs)
   Q_Measure = Q_Measure - 1
   
   'cargamos el combo con los datos seleccionados
   For I = 0 To Q_Measure
          CbnMeasure.AddItem ListInt(1, I)
   Next
   
   GridList_Input.Visible = True
End Sub

'para cargar el combo de medidas segun el insumo
Private Sub CbnMeasure_lostFocus()

   Dim C_Proc As New C_General_Procedures
   Dim C_ListInt As New C_List_Price
   
   Dim IndexInputs As String
   Dim IndexMeasure As String
   
   Dim id As Integer
     
   Dim ListInt() As Variant
     
   IndexInputs = CbnImputs.ListIndex
   
   On Error GoTo ctrlerr
    
   'capturamos el id de la medida
   id = C_Proc.Recover_Id_Detail("TC_Measure", CbnMeasure.Text, IndexInputs)
   
   'cargamos consulta datos medidas según el insumo solicitado en BD
   ListInt = C_ListInt.Charge_List_View_sale(IndexInputs, id)
   
   'traemos la cantidad de medidas según el insumo solicitado en BD
   Q_Charge_List_View = C_ListInt.Q_Charge_List_View_sale(IndexInputs, id)
   Q_Charge_List_View = Q_Charge_List_View - 1
   
   'dimencionamos el grid
   GridList_Input.Rows = Q_Charge_List_View + 2
   
   'inicializamos variables
   IFF = 1
   Columnas = 4
   
   'cargamos el array
   For I = 0 To Q_Charge_List_View
      For IC = 0 To Columnas
          GridList_Input.TextMatrix(IFF, IC) = ListInt(IC, I)
      Next
      IFF = IFF + 1
   Next
   
   IFF = 1
    
   'CREAR BOTON SELECCIONAR
   For I = 0 To Q_Charge_List_View
        GridList_Input.TextMatrix(IFF, 5) = "INGRESAR"
        IFF = IFF + 1
   Next
     
   'redimencionamos el tamaño de las columnas a los datos digitados
    For Row = 0 To GridList_Input.Rows - 1
        For Col = 0 To GridList_Input.Cols - 1
            GridList_Input.ColWidth(Col) = IIf(Me.TextWidth(GridList_Input.TextMatrix(Row, Col)) + 400 > GridList_Input.ColWidth(Col), Me.TextWidth(GridList_Input.TextMatrix(Row, Col)) + 400, GridList_Input.ColWidth(Col))
        Next
    Next
    
   Exit Sub
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case 13
    Exit Sub
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    
End Sub

'se realiza operaciones para saber el valor del material para el grid de operacion
Private Sub GridList_Input_Click()
    
    Dim PosGrid() As Variant
    Dim id_GInput As String
    
    Dim Price As Long
    Dim Price_Total As String
    Dim Description As String
    Dim Rows_Q As Integer
    Dim Q_Inputs As Long
    
    
    If TxtQuanty.Text = "" Then
        MsgBox "El campo cantidad debe estar diligenciado!", vbExclamation + vbOKOnly, "Información!"
        TxtQuanty.SetFocus
        
    Else
         Q_Inputs = TxtQuanty.Text
         id_GInput = GridList_Input.Row
         
         Description = GridList_Input.TextMatrix(id_GInput, 3)
         Price = GridList_Input.TextMatrix(id_GInput, 4)
            
         Price_Total = Price * Q_Inputs
            
         Columnas = 2
          
         Rows_Q = GridList_Operative.Rows
         Rows_Q = Rows_Q + 1
         
         GridList_Operative.Rows = Rows_Q
         
         
        'cargamos el array
        GridList_Operative.TextMatrix(Count_GOperative, 0) = Description
        GridList_Operative.TextMatrix(Count_GOperative, 1) = Price
        GridList_Operative.TextMatrix(Count_GOperative, 2) = Q_Inputs
        GridList_Operative.TextMatrix(Count_GOperative, 3) = Price_Total
        GridList_Operative.TextMatrix(Count_GOperative, 4) = "Eliminar"
        
        Count_GOperative = Count_GOperative + 1
        
        'redimencionamos el tamaño de las columnas a los datos digitados
         For Row = 0 To GridList_Operative.Rows - 1
             For Col = 0 To GridList_Operative.Cols - 1
                 GridList_Operative.ColWidth(Col) = IIf(Me.TextWidth(GridList_Operative.TextMatrix(Row, Col)) + 400 > GridList_Operative.ColWidth(Col), Me.TextWidth(GridList_Operative.TextMatrix(Row, Col)) + 400, GridList_Operative.ColWidth(Col))
             Next
         Next
             
         FrmOperative.Visible = True
         BtnCreate.Visible = True
         Sum_Values (Price_Total)
    
    End If
   
End Sub

'operaciones de eliminacion de datos del grid de seleccion de insumos
Private Sub GridList_Operative_Click()

    Dim id_GInput As String
    Dim Price_Total As Long
    Dim Price_Rest As Long
    Dim Total_Result As Long
     
    On Error GoTo ctrlerr
    
    If MsgBox("Esta seguro de eliminar el insumo??", vbYesNo, "Confirmacion") = vbYes Then
        
        'Acciones a realizar
        id_GInput = GridList_Operative.Row
        Price_Rest = GridList_Operative.TextMatrix(id_GInput, 3)
        Price_Total = LblValue_Neto.Caption
        Total_Result = Price_Total - Price_Rest
        MsgBox Total_Result
        
        GValueNeto = 0
        GValueDouble = 0
        GValueWinner = 0
        GValueSubTotal = 0
        GValueIva = 0
        GValueTotal = 0
        GVAbono = 0
               
        GSwitch_Delete = 1
               
        Sum_Values (Total_Result)
        
        GridList_Operative.TextMatrix(id_GInput, 0) = ""
        GridList_Operative.TextMatrix(id_GInput, 1) = ""
        GridList_Operative.TextMatrix(id_GInput, 2) = ""
        GridList_Operative.TextMatrix(id_GInput, 3) = ""
        GridList_Operative.TextMatrix(id_GInput, 4) = ""
        
    End If
    
    Exit Sub
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case 13
     MsgBox "El insumo fue eliminado!", vbExclamation + vbOKOnly, "Información!"
    Exit Sub
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
     
End Sub

'OPCION PARA BUSCAR POR DOCUMENTO
Private Sub Opdoc_Click()
    
    Crear_client = 0
    
    Dim C_Proc As New C_General_Procedures
    Dim cargar_datos() As Variant
    
    CbnSearch.clear
    CbnSearch.Text = "Seleccione..."
    CbnSearch.Width = 2500
    
    'cargamos consulta datos documentos en BD
    cargar_datos = C_Proc.Datos_Charge("Client", "Add_Doc")
    
    'traemos la cantidad de doc en BD
    Q_cargar_datos = C_Proc.Q_Combo("Client")
    Q_cargar_datos = Q_cargar_datos - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_cargar_datos
        CbnSearch.AddItem cargar_datos(1, I)
    Next

End Sub

'OPCION PARA BUSCAR POR NOMBRE
Private Sub OpName_Click()
      
    Crear_client = 0
      
    Dim C_Proc As New C_General_Procedures
    Dim cargar_datos() As Variant
    
    CbnSearch.clear
    CbnSearch.Text = "Seleccione..."
    CbnSearch.Width = 5865
    
    'cargamos consulta datos documentos en BD
    cargar_datos = C_Proc.Datos_Charge("Client", "Add_Name")
    
    'traemos la cantidad de doc en BD
    Q_cargar_datos = C_Proc.Q_Combo("Client")
    Q_cargar_datos = Q_cargar_datos - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_cargar_datos
        CbnSearch.AddItem cargar_datos(1, I)
    Next

End Sub
'VALIDAR CAMPO NUMERICO y consultar el nuevo saldo
Private Sub TxtCast_Change()

    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    Dim ResultSald As Long
    Dim Operator As Long
    
    initial = TxtCast.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtCast.Text = final
    
    If TxtCast.Text = "" Then
        Operator = 0
    Else
        Operator = TxtCast.Text
    End If
    
    ResultSald = Rest_Sald(Format(LblValue_Total.Caption, "##"), Operator)
    LblValue_Sald.Caption = ResultSald
    LblValue_Sald.Caption = Format(LblValue_Sald.Caption, "####,####")
End Sub

'VALIDAR CAMPO NUMERICO y asignar fechafinal al calendario
Private Sub TxtDaysEnd_Change()
    
    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    Dim ValueDate As Integer
    Dim DateEnd As Date
    Dim DateInitial As Date
        
    initial = TxtDaysEnd.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtDaysEnd.Text = final

    If TxtDaysEnd.Text = "" Then
        ValueDate = 0
    Else
        ValueDate = TxtDaysEnd.Text
    End If
        
    DPSale.Value = DateAdd("y", ValueDate, Date)
    
End Sub

'VALIDAR CAMPO NUMERICO
Private Sub TxtQuanty_Change()
    
    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    initial = TxtQuanty.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtQuanty.Text = final

End Sub
''''''----------- END REGION EVENTOS

''''''----------- REGION FUNCIONES
'VALIDA CAMPOS OBLIGATORIOS
Function ValidateCampos(verificar As Integer) As Integer
    
    'instanciamos variables
    Dim validate As Integer
      
    'inicializamos en 0
    validate = 0
    
    Select Case verificar
    
        Case 1
        
            If CbnSearch.Text = "Seleccione..." Then
                
                MsgBox "no se ha seleccionado Cliente!!!", vbExclamation + vbOKOnly, "Información!"
                validate = 1
              
            End If
          
        Case 2
        
        If GType_operation <> "Factura" Then
             
             If CbnSearch.Text = "Seleccione..." Or TxtDescripProject = "" Or TxtDaysEnd.Text = "" Then
                
                If CbnSearch.Text = "Seleccione..." Then
                    CbnSearch.BackColor = &H80&
                Else
                    CbnSearch.BackColor = &H80000005
                End If
                If TxtDescripProject.Text = "" Then
                    TxtDescripProject.BackColor = &H80&
                Else
                    TxtDescripProject.BackColor = &H80000005
                End If
                 If TxtDaysEnd.Text = "" Then
                    TxtDaysEnd.BackColor = &H80&
                Else
                    TxtDaysEnd.BackColor = &H80000005
                End If

                validate = 1
             
             End If
             
        Else
             
             If CbnSearch.Text = "Seleccione..." Or TxtDescripProject = "" Or TxtDaysEnd.Text = "" Or TxtCast.Text = "" Then
                
                If CbnSearch.Text = "Seleccione..." Then
                    CbnSearch.BackColor = &H80&
                Else
                    CbnSearch.BackColor = &H80000005
                End If
                If TxtDescripProject.Text = "" Then
                    TxtDescripProject.BackColor = &H80&
                Else
                    TxtDescripProject.BackColor = &H80000005
                End If
                If TxtDaysEnd.Text = "" Then
                    TxtDaysEnd.BackColor = &H80&
                Else
                    TxtDaysEnd.BackColor = &H80000005
                End If
                If TxtCast.Text = "" Then
                    TxtCast.BackColor = &H80&
                Else
                    TxtCast.BackColor = &H80000005
                End If
             
                validate = 1
                
             End If
        
        End If
            
        Case Else
        
    End Select
    
      ValidateCampos = validate
          

End Function

'BLOQUEAR CAMPOS
Function block()

    TxtDocument.Enabled = False
    TxtTypeDocument.Enabled = False
    TxtAddress.Enabled = False
    TxtPhone.Enabled = False
    TxtEmail.Enabled = False
    TxtObservations.Enabled = False
  
End Function

'SUMAR TODOS LOS VALORES
Function Sum_Values(Value As Long)
   
    'averiguamos si es el primer valor
    If LblValue_Neto.Caption = "" Then
        GValueNeto = 0
    Else
        GValueNeto = Format(LblValue_Neto.Caption, "##")
    End If
     
    'validamos si borraron un dato
    If GSwitch_Delete = 1 Then
        GValueNeto = 0
    End If
     
    'sumamos y asignamos el valor de los materiales
    GValueNeto = GValueNeto + Value
    LblValue_Neto.Caption = GValueNeto
    LblValue_Neto.Caption = Format(LblValue_Neto.Caption, "####,####")
     
    'multiplicamos y asignamos el valor de la mano de obra
    GValueDouble = GValueNeto * 2
    LdValue_Double.Caption = GValueDouble
    LdValue_Double.Caption = Format(LdValue_Double.Caption, "####,####")
     
    'multiplicamos y asignamos el valor de la mano de obra
    GValueWinner = GValueNeto * 0.4
    LblValue_Winner.Caption = GValueWinner
    LblValue_Winner.Caption = Format(LblValue_Winner.Caption, "####,####")
    
    'sumamos subtotales
    GValueSubTotal = GValueNeto + GValueDouble + GValueWinner
    LblValue_Subtotal.Caption = GValueSubTotal
    LblValue_Subtotal.Caption = Format(LblValue_Subtotal.Caption, "####,####")
    Dim OperateIva As String
    
    'traer valor de operacion iva
    OperateIva = "16"
    
    'validamos si requiere iva
    If OpYes = True Then
        
        GValueIva = (GValueSubTotal * OperateIva) / 100
        GValueTotal = GValueSubTotal + GValueIva
        LblValue_Iva.Caption = GValueIva
        LblValue_Iva.Caption = Format(LblValue_Iva.Caption, "####,####")
        LblValue_Total.Caption = GValueTotal
        LblValue_Total.Caption = Format(LblValue_Total.Caption, "####,####")
       
    Else
        
        GValueTotal = GValueSubTotal
        LblValue_Iva.Caption = 0
        LblValue_Total.Caption = GValueTotal
        LblValue_Total.Caption = Format(LblValue_Total.Caption, "####,####")
    
    End If
    
    'calculamos el abono sugerido
    GVAbono = GValueTotal / 2
    LblValue_CastSugery.Caption = GVAbono
    LblValue_CastSugery.Caption = Format(LblValue_CastSugery.Caption, "####,####")
    
    TxtQuanty.Text = ""
    GSwitch_Delete = 0
 End Function

'AVERIGUAR SALDO
Function Rest_Sald(minuendo As Long, Operator As Long)

    Dim Result As Long
    Result = minuendo - Operator
    Rest_Sald = Result

End Function
''''''-----------END  REGION FUNCIONES

''''''----------- REGION FUNCIONES BD
'GENERAR FACTURA
Function G_Factura()
    
    Dim op_Search As String
    Dim id As Integer
    Dim Id_User As Integer
    Dim GUARDAR  As String
    Dim C_Proc As New C_General_Procedures
    Dim C_Project As New C_Project
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
    
    'capturamos el cliente de la operacion
    id = C_Proc.Recover_Id(op_Search, "Client", CbnSearch.Text)
    Id_User = C_Proc.Recover_Id("User", "Users", MenuCarpenter.Lbl_Value_User.Caption)
    GUARDAR = C_Project.Add_Project(GType_operation, LblNumber.Caption, id, TxtDescripProject.Text, LblValue_Date.Caption, TxtDaysEnd.Text, Format(LblValue_Neto.Caption, "##"), Format(LdValue_Double.Caption, "##"), Format(LblValue_Winner.Caption, "##"), Format(LblValue_Subtotal.Caption, "##"), Format(LblValue_Iva.Caption, "##"), Format(LblValue_Total.Caption, "##"), TxtCast.Text, Format(LblValue_Sald.Caption, "##"), Id_User)
    
    'validamos el resultado de la operacion anterior
    If GUARDAR = "OK" Then
    
        export_factura
        export_CCobro
        
        LblhelpGeneral.Caption = GType_operation & " realizada con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        
    Else
        
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If
    
End Function

'EXPORTAR A EXCEL FACTURA
Function export_factura()

    Dim xlsApp As Object
    Dim xlsBook As Object
    Dim StringXL As String
    Dim GUARDAR As String
     
    LblhelpGeneral.Visible = True
    LblhelpGeneral.Caption = "Generando factura...."
    LblhelpGeneral.ForeColor = &H8000&
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.ScreenUpdating = True
    
    StringXL = App.Path & "\Factura.xlsx"
    
    Set xlsBook = xlsApp.Workbooks.Open(StringXL)
    
    Dim xlsSheet As Object
    Set xlsSheet = xlsBook.WorkSheets("Hoja1")
    
    xlsSheet.Range("H4") = LblValue_Date.Caption 'fecha de creacion de la factura
    xlsSheet.Range("N3") = LblNumber.Caption 'N de factura
    xlsSheet.Range("H6") = DPSale.Value 'fecha de finalizadion de la factura
    xlsSheet.Range("M5") = "N/A" 'orden de compra
    xlsSheet.Range("M6") = "Efectivo" 'forma de pago
    xlsSheet.Range("C7") = GName 'nombre cliente
    xlsSheet.Range("M7") = TxtDocument.Text 'documento
    xlsSheet.Range("C8") = TxtAddress.Text 'direccion
    xlsSheet.Range("H8") = TxtPhone.Text 'telefono
    xlsSheet.Range("M8") = "Bogotá" 'telefono
    xlsSheet.Range("B11") = TxtDescripProject.Text 'descripcion
    xlsSheet.Range("L11") = Format(LblValue_Subtotal.Caption, "##") 'vr unitario
    xlsSheet.Range("N11") = Format(LblValue_Subtotal.Caption, "##") 'vr total uni
    xlsSheet.Range("N22") = Format(LblValue_Subtotal.Caption, "##") 'vr sub total
    xlsSheet.Range("N24") = Format(LblValue_Iva.Caption, "##") 'vr iva
    xlsSheet.Range("N25") = Format(LblValue_Total.Caption, "##") 'vr total
    xlsSheet.Range("N27") = TxtCast.Text 'vr abono
    xlsSheet.Range("N28") = Format(LblValue_Sald.Caption, "##") 'vr saldo
    
    RUTA = App.Path & "\FACTURAS"
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' Se crea la instancia
    
    If fso.FolderExists(RUTA) Then
      Else
           fso.CreateFolder (RUTA)
    End If
    
    GUARDAR = RUTA & "\"
       
    xlsBook.SaveAs GUARDAR & "FACTURA_" & LblNumber.Caption & ".XLSX"
    xlsBook.Close (StringXL)
        
End Function

Function export_CCobro()

    Dim xlsApp As Object
    Dim xlsBook As Object
    Dim StringXL As String
    Dim GUARDAR As String
    
    LblhelpGeneral.Caption = "Generando Cuenta de cobro...."
    LblhelpGeneral.ForeColor = &H8000&
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.ScreenUpdating = True
    
    StringXL = App.Path & "\C_Cobro.xlsx"
    
    Set xlsBook = xlsApp.Workbooks.Open(StringXL)
    
    Dim xlsSheet As Object
    Set xlsSheet = xlsBook.WorkSheets("Hoja1")
    
    xlsSheet.Range("B7") = LblValue_Date.Caption 'fecha de creacion de la factura
    xlsSheet.Range("H7") = LblNumber.Caption & "_01" 'N de cotizacion
    xlsSheet.Range("A10") = GName 'nombre cliente
    xlsSheet.Range("A11") = TxtDocument.Text 'documento
    xlsSheet.Range("A12") = TxtPhone.Text 'telefono
    xlsSheet.Range("G18") = Format(LblValue_Total.Caption, "##") 'vr total
    xlsSheet.Range("A21") = TxtDescripProject.Text 'descripcion
        
    RUTA = App.Path & "\C_COBROS"
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' Se crea la instancia
    
    If fso.FolderExists(RUTA) Then
      Else
           fso.CreateFolder (RUTA)
    End If
    
    GUARDAR = RUTA & "\"
       
    xlsBook.SaveAs GUARDAR & "C_Cobro_" & LblNumber.Caption & "_N1" & ".XLSX"
    xlsBook.Close (StringXL)

End Function
        

'GENERAR COTIZACION
Function G_Cotizacion()
    
    Dim op_Search As String
    Dim id As Integer
    Dim Id_User As Integer
    Dim GUARDAR  As String
    Dim C_Proc As New C_General_Procedures
    Dim C_Project As New C_Project
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
    
    'capturamos el cliente de la operacion
    id = C_Proc.Recover_Id(op_Search, "Client", CbnSearch.Text)
    Id_User = C_Proc.Recover_Id("User", "Users", MenuCarpenter.Lbl_Value_User.Caption)
    GUARDAR = C_Project.Add_Project(GType_operation, LblNumber.Caption, id, TxtDescripProject.Text, LblValue_Date.Caption, TxtDaysEnd.Text, Format(LblValue_Neto.Caption, "##"), Format(LdValue_Double.Caption, "##"), Format(LblValue_Winner.Caption, "##"), Format(LblValue_Subtotal.Caption, "##"), Format(LblValue_Iva.Caption, "##"), Format(LblValue_Total.Caption, "##"), 0, 0, Id_User)
    
    'validamos el resultado de la operacion anterior
    If GUARDAR = "OK" Then
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = GType_operation & " realizada con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        export_cotizacion
    
    Else
        
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If
    
End Function

'EXPORTAR A EXCEL COTIZACION
Function export_cotizacion()

    Dim xlsApp As Object
    Dim xlsBook As Object
    Dim StringXL As String
    Dim GUARDAR As String
  
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.ScreenUpdating = True
    
    StringXL = App.Path & "\Cotizacion.xlsx"
    
    Set xlsBook = xlsApp.Workbooks.Open(StringXL)
    
    Dim xlsSheet As Object
    Set xlsSheet = xlsBook.WorkSheets("Hoja1")
    
    xlsSheet.Range("B7") = LblValue_Date.Caption 'fecha de creacion de la factura
    xlsSheet.Range("H7") = LblNumber.Caption 'N de cotizacion
    xlsSheet.Range("E30") = TxtDaysEnd.Text 'fecha de finalizadion de la factura
    xlsSheet.Range("A10") = "Sr(es). " & GName 'nombre cliente
    xlsSheet.Range("A15") = TxtDescripProject.Text 'descripcion
    xlsSheet.Range("G23") = Format(LblValue_Total.Caption, "##") 'vr total
    
    If OpYes = True Then
        xlsSheet.Range("A26") = "Los precios especificados ya tiene el incluyen i.v.a. e impuestos a la Industria" 'con iva
    Else
        xlsSheet.Range("A26") = "Los precios especificados no incluyen i.v.a. ni impuestos a la Industria" 'sin iva
    End If
    
    RUTA = App.Path & "\COTIZACIONES"
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' Se crea la instancia
    
    If fso.FolderExists(RUTA) Then
      Else
           fso.CreateFolder (RUTA)
    End If
    
    GUARDAR = RUTA & "\"
       
    xlsBook.SaveAs GUARDAR & "COTIZACION_" & LblNumber.Caption & ".XLSX"
    xlsBook.Close (StringXL)
        
End Function

''''''-----------END REGION FUNCIONES BD


