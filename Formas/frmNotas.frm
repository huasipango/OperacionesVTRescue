VERSION 5.00
Begin VB.Form frmNotas 
   Caption         =   "Notas de Credito"
   ClientHeight    =   7905
   ClientLeft      =   2040
   ClientTop       =   2205
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Observaciones"
      Height          =   735
      Left            =   240
      TabIndex        =   61
      Top             =   6000
      Width           =   7815
      Begin VB.TextBox txtobs 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   62
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.ListBox lstPlazas 
      Height          =   1815
      Left            =   7920
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame frBotones 
      Height          =   855
      Left            =   240
      TabIndex        =   23
      Top             =   6840
      Width           =   7815
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualiza Folios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   525
         Left            =   6720
         Picture         =   "frmNotas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   525
         Left            =   5880
         Picture         =   "frmNotas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frConcepto 
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   4095
      Begin VB.OptionButton optConc 
         Caption         =   "Bonificación sobre Ventas"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   3615
      End
      Begin VB.OptionButton optConc 
         Caption         =   "Bonificación especial por comision mal aplicada"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   37
         Top             =   490
         Width           =   3735
      End
      Begin VB.OptionButton optConc 
         Caption         =   "Cancelación de Factura"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   800
         Width           =   3495
      End
      Begin VB.OptionButton optConc 
         Caption         =   "Dispersion devuelta segun detalle anexo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame frProd 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   4095
      Begin VB.ComboBox cboProductos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   63
         Text            =   "cboProductos"
         Top             =   360
         Width           =   3735
      End
      Begin VB.OptionButton optProd 
         Caption         =   "Bono Elect."
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optProd 
         Caption         =   "Despensa"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optProd 
         Caption         =   "Gasolina"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame frCte 
      Caption         =   "NOTAS DE CREDITO "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   7815
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   7575
         Begin VB.TextBox txtPlaza 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Height          =   255
            Left            =   2400
            TabIndex        =   59
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox chkOI 
            Caption         =   "Otros Ingresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4080
            TabIndex        =   58
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtPedido 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5160
            TabIndex        =   57
            Top             =   1560
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboBodegas 
            Height          =   315
            ItemData        =   "frmNotas.frx":0204
            Left            =   1200
            List            =   "frmNotas.frx":020B
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtCliente 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdBuscarC 
            BackColor       =   &H00C0C0C0&
            CausesValidation=   0   'False
            Height          =   375
            Left            =   2400
            Picture         =   "frmNotas.frx":022D
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtFecha 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   47
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtFactura 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   1150
            Width           =   1095
         End
         Begin VB.TextBox txtSerie 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   46
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox txtFolio 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   45
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Plaza:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Sucursal:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   53
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Factura:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1150
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Serie:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   51
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Sig. Folio:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   50
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblNombre 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   49
            Top             =   720
            Width           =   4455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   4215
         Begin VB.CommandButton cmdAbrir 
            BackColor       =   &H00C0C0C0&
            CausesValidation=   0   'False
            Height          =   450
            Left            =   3480
            Picture         =   "frmNotas.frx":032F
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Subir desde archivo"
            Top             =   160
            Width           =   450
         End
         Begin VB.OptionButton optLectura 
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optLectura 
            Caption         =   "Ajuste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtRem 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lblPlaza 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame frGrav 
      Height          =   2655
      Left            =   4440
      TabIndex        =   24
      Top             =   3240
      Width           =   3615
      Begin VB.TextBox txtTotalGra 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtIvaComGra 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtComGra 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtIvaGra 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtGra 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "T O T A L "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "IVA Comisión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Comisión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "IVA Bonos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Bonos Gravados:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frExentos 
      Height          =   2535
      Left            =   4440
      TabIndex        =   16
      Top             =   3240
      Width           =   3615
      Begin VB.TextBox txtExe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtComExe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtIvaComExe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTotalExe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bonos Exentos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Comisión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "IVA Comisión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "T O T A L "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumeroNota As Integer
Dim ncNormal_OI As Integer
Private Sub cboBodegas_Click()
    If chkOI.value = 1 Then
      sqls = " select Prefijo as serie from FOLIOS"
      sqls = sqls & " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
      sqls = sqls & " AND TIPO='FCM'"
    Else
      sqls = " select serie_factura as serie from bodegas"
      sqls = sqls & " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    End If
        
    Set rsbod = New ADODB.Recordset
    rsbod.Open sqls, cnxBD, adOpenForwardOnly
    
    If Not rsbod.EOF Then
        txtserie = rsbod!serie
        BuscaSigFolio (cboBodegas.ItemData(cboBodegas.ListIndex))
        Call STAT
    Else
        MsgBox "Error en la serie de la sucursal"
    End If
End Sub

Private Sub chkOI_Click()
  If chkOI.value = 0 Then
   ncNormal_OI = 3
  Else
   ncNormal_OI = 5
  End If
End Sub

Private Sub cmdAbrir_Click()
Dim Totally As Double

If Trim(txtFactura) = "" Then
    MsgBox "Debe capturar el número de factura al cual se le va a aplicar la nota de crédito", vbInformation
    Exit Sub
End If

If Trim(txtCliente.Text) = "" Then
    MsgBox "Debe capturar el número de cliente al cual se le va a aplicar la nota de crédito", vbInformation
    Exit Sub
End If

If Trim(txtRem.Text) = "" Then
    MsgBox "Debe capturar el número de entrada o ajuste", vbInformation
    Exit Sub
End If


sql = " select bon_fac_pedido from bon_factura"
sql = sql & " Where bon_fac_sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex)
sql = sql & " and bon_fac_numero  = " & Val(txtFactura)
sql = sql & " and bon_fac_serie = '" & Trim(txtserie) & "'"
sql = sql & " and bon_fac_cliente  = " & Val(txtCliente) & ""


Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
    pedido = rsBD!bon_fac_pedido
End If
rsBD.Close

sql = " sp_FactBEVarios"
sql = sql & " @sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex)
sql = sql & ", @pedido  = " & pedido
sql = sql & ", @accion = 'BuscaFacturaConc'"

Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    Producto = rsBD!bon_fac_tpobon
    If Producto >= 1 Then
       Totally = rsBD!bon_fac_bonexe + rsBD!bon_fac_bongra + rsBD!bon_fac_ivagra + rsBD!bon_fac_comision + rsBD!ivacomision
    End If
    txtPedido.Text = pedido
    valor = IIf(Producto <> 1, rsBD!bon_fac_bongra + rsBD!bon_fac_ivagra, rsBD!bon_fac_bonexe)
    valorgra = rsBD!bon_fac_bongra
Else
    MsgBox "La factura no existe o no corresponde al cliente", vbCritical
    Exit Sub
End If

txtobs = ""

'-------------
optConc(0).Enabled = False
optConc(1).Enabled = False
optConc(2).Enabled = False
optConc(3).Enabled = False


If optLectura(0).value = True Then   'Por entrada

    If Val(txtRem.Text) = 0 Then     '*******Sin lectura*********
        If pedido = 0 Then 'reposicion
            optConc(1).Enabled = False   'Para entradas sin lectura solo cancelación de
            optConc(2).Enabled = True
            optConc(3).Enabled = False
            optConc(2).value = True
            
            sql = " select * from facturasreposiciones" & _
                  " Where Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                  " and factura = " & Val(txtFactura)
            Set rsBD2 = New ADODB.Recordset
            rsBD2.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If rsBD2.EOF Then 'No es reposicion de factura, sino factura de prefactura
               Call LeeproductoBE(cboProductos, "sp_sel_productobe '" & Producto & "','','Nprod'")
               cboProductos.Text = UCase(texto_producto)
               If Product = 1 Then
                  frExentos.Visible = True
                  frGrav.Visible = False
               Else
                  frExentos.Visible = False
                  frGrav.Visible = True
               End If
               
               optConc(1).Enabled = True
               optConc(1).value = True
            '   optConc(2).Enabled = False
            
               frGrav.Enabled = True
            End If
               
           If Producto = 1 And optConc(1).value = False Then
               'optProd(0).Value = True
               Call LeeproductoBE(cboProductos, "sp_sel_productobe '" & Producto & "','','Nprod'")
               cboProductos.Text = UCase(texto_producto)
               If Product = 1 Then
                  frExentos.Visible = True
                  frGrav.Visible = False
               Else
                  frExentos.Visible = False
                  frGrav.Visible = True
               End If
               txtExe.Text = Format(rsBD!bon_fac_bonexe, "######0.00")
               txtComExe.Text = Format(rsBD!bon_fac_comision, "######0.00")
               txtIvaComExe.Text = Format(rsBD!ivacomision, "######0.00")
           ElseIf Producto = 2 Then
               'optProd(1).Value = True
               Call LeeproductoBE(cboProductos, "sp_sel_productobe '" & Producto & "','','Nprod'")
               cboProductos.Text = UCase(texto_producto)
               If Product = 1 Or Product = 6 Then
                  frExentos.Visible = True
                  frGrav.Visible = False
               Else
                  frExentos.Visible = False
                  frGrav.Visible = True
               End If
               txtGra.Text = Format(rsBD!bon_fac_bongra, "######0.00")
               txtIvaGra.Text = Format(rsBD!bon_fac_ivagra, "######0.00")
               txtComGra.Text = Format(rsBD!bon_fac_comision, "######0.00")
               txtIvaComGra.Text = Format(rsBD!ivacomision, "######0.00")
           End If
            frBotones.Enabled = True
            Frame3.Enabled = True
            
        Else
            optConc(1).Enabled = True   'Para entradas sin lectura solo cancelación de
            'optConc(2).Enabled = True
          '  optConc(3).Enabled = True
         '   optConc(2).Enabled = True
            optConc(2).value = True
            frConcepto.Enabled = True
            Call LeeproductoBE(cboProductos, "sp_sel_productobe '" & Producto & "','','Nprod'")
            cboProductos.Text = UCase(texto_producto)
            '--VALIDAR SI YA AY NOTA DE CREDITO
               sqls = "SELECT SUM(ISNULL(VALOR,0))Valor FROM NOTASCRE WITH (NOLOCK)"
               sqls = sqls & " WHERE BODEGA=" & cboBodegas.ItemData(cboBodegas.ListIndex)
               sqls = sqls & " AND CLIENTE=" & Val(txtCliente)
               sqls = sqls & " AND Factura=" & Val(txtFactura)
               sqls = sqls & " AND SerieFac='" & Trim(txtserie) & "'"
               sqls = sqls & " AND STATUS<>2"
               Set rsbd4 = New ADODB.Recordset
               rsbd4.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
               If Not rsbd4.EOF Then
                  If (Totally - rsbd4!valor) = 0 Then
                      MsgBox "Ya existe una nota de credito que cubre la totalidad de la factura", vbCritical, "Ya tiene nota de credito"
                      Exit Sub
                  End If
               End If
            '------

            If Producto >= 1 Then
               optConc(2).Enabled = True
               
               If Product = 1 Then
                  frExentos.Visible = True
                  frGrav.Visible = False
                  '-----------------------------
                  txtExe.Text = Format(rsBD!bon_fac_bonexe, "######0.00")
                  txtComExe.Text = Format(rsBD!bon_fac_comision, "######0.00")
                  txtIvaComExe.Text = Format(rsBD!ivacomision, "######0.00")
               Else
                  frExentos.Visible = False
                  frGrav.Visible = True
                  '-----------------------
                  txtGra.Text = Format(rsBD!bon_fac_bongra, "######0.00")
                  txtIvaGra.Text = Format(rsBD!bon_fac_ivagra, "######0.00")
                  txtComGra.Text = Format(rsBD!bon_fac_comision, "######0.00")
                  txtIvaComGra.Text = Format(rsBD!ivacomision, "######0.00")
               End If
            End If
            
            frBotones.Enabled = True
            Frame3.Enabled = True
        End If
    
    End If
    
ElseIf optLectura(1).value = True Then 'Por ajuste
    If Val(txtRem.Text) = 0 Then
        MsgBox "El ajuste no puede ser cero", vbInformation
        txtRem.Text = ""
        LimpiafrExe
        LimpiafrGrav
        frExentos.Visible = False
        frGrav.Visible = False
        Exit Sub
    Else
            
            sql = " select reembolso from notascre"
            sql = sql & " Where reembolso = " & Val(txtRem.Text)
            Set rsBD2 = New ADODB.Recordset
            rsBD2.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

            If Not rsBD2.EOF Then
               MsgBox "Ya existe una nota de credito con ese numero de AJUSTE, por favor verifiquelo", vbCritical, "Notas de crédito"
               txtRem = ""
              Exit Sub
            End If
            
            Call LeeproductoBE(cboProductos, "sp_sel_productobe '" & Producto & "','','Nprod'")
            cboProductos.Text = UCase(texto_producto)
            If Product = 1 Then
                  frExentos.Visible = True
                  frGrav.Visible = False
            Else
                  frExentos.Visible = False
                  frGrav.Visible = True
            End If
            
            frConcepto.Enabled = True
            optConc(0).Enabled = True  'habilitado, solo para BE
            optConc(0).value = True
            optConc(1).Enabled = False
            optConc(2).Enabled = False   'se habilita solo para BE
            optConc(3).Enabled = False   ' se habilita si se tecleea un cliente valido como la uni
            
            sqls = "select * from ajustesbe" & _
                " Where folio = " & txtRem & _
                " and concepto IN (2,4,5,7,8,9)" & _
                " and cliente = " & txtCliente
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
                If Product = 1 Then
                   txtExe.Text = Format(rsBD!cargo, "######0.00")
                   txtComExe.Text = Format(0, "######0.00")
                   txtIvaComExe.Text = Format(0, "######0.00")
                Else
                   
                   Totala = Format(rsBD!cargo, "######0.00")
                   sqls = "select impuestointerior iva from bodegas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
                   Set rsBD2 = New ADODB.Recordset
                   rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
               
                   If Not rsBD.EOF Then
                      TotalGra = Round(Totala / (1 + (rsBD2!iva / 100)), 2)
                      ivagra = Totala - TotalGra
                   Else
                       TotalGra = Round(Totala / (1 + (16 / 100)), 2)
                       ivagra = Totala - TotalGra
                   End If
                   txtGra.Text = Format(TotalGra, "######0.00")
                   txtIvaGra.Text = Format(ivagra, "######0.00")
                   txtTotalGra = Format(rsBD!cargo, "######0.00")
                End If
                frBotones.Enabled = True
                Frame3.Enabled = True
            Else
                MsgBox "Ajuste inválido, no existe o no corresponde al cliente", vbCritical
                txtRem.Text = ""
                Exit Sub
            End If
    End If
End If

If gnBodega <> 1 Then 'solamente en monterrey se pueden hacer notas de credito por bonificacion x comision mal aplicada
    optConc(1).Enabled = False
End If

End Sub

Private Sub cmdActualiza_Click()
Dim resp As Integer
Dim Folio As Integer
Dim Bodega As Integer

Bodega = cboBodegas.ItemData(cboBodegas.ListIndex)

sqls = " select Prefijo serie, consecutivo " & _
      " From folios " & _
      " Where Bodega =" & Bodega & _
      " and tipo = 'NC'"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockPessimistic

   If Not rsBD.EOF Then
      Folio = rsBD!consecutivo
   Else
      MsgBox "Error en folio de la Factura, verifique con sistemas", vbCritical
      Exit Sub
   End If
      
On Error GoTo err_panel
   resp = InputBox("Ultimo folio de Notas", "Folio de Facturas de Comisiones", Folio)
   
On Error GoTo err_gral

   sqls = "update folios set consecutivo = " & resp & _
          " , fechamodificacion = getdate()" & _
          " Where Bodega =" & Bodega & _
          " and tipo = 'NC'"
   
   cnxBD.Execute sqls, intRegistros
   
   Exit Sub
   
err_panel:
         Exit Sub
err_gral:
         MsgBox ERR.Description
         Exit Sub
   
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtCliente = frmConsulta.cliente
       lblNombre = frmConsulta.Nombre
    End If
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub
Private Sub cmdGrabar_Click()
Dim TIPO_MOV, Concepto As Integer
Dim importe As Double, Comision As Double
Dim TipoBon As Integer
Dim serieFac As String
Dim serif As String
On Error GoTo err_gral
       
    sql = " select bon_fac_serie ,bon_fac_bonexe, bon_fac_bongra, bon_fac_ivagra, bon_fac_comision, "
    sql = sql & " bon_fac_ivacomis as ivacomision, bon_fac_tpobon, bon_fac_pedido PEDIDO , bon_Fac_status from bon_factura"
    sql = sql & " Where bon_fac_sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    sql = sql & " and bon_fac_numero  = " & Val(txtFactura)
    sql = sql & " and bon_fac_serie = '" & Trim(txtserie) & "'"
    sql = sql & " and bon_fac_cliente  = " & Val(txtCliente) & ""
        
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        
        If rsBD!bon_fac_tpobon >= 1 Then
           tpb = rsBD!bon_fac_tpobon
           Product = tpb
        End If

        'If optProd(0).Value = True Or optProd(2).Value = True Then
        If Product = 1 Then
            If Val(txtComExe.Text) > rsBD!bon_fac_comision Then
                MsgBox "La comision de la nota no puede ser mayor a la comision de la factura", vbCritical
                Exit Sub
            End If
        Else ' optProd(1).Value = True Then 'Gasolina
            If Val(txtComGra.Text) > rsBD!bon_fac_comision Then
                MsgBox "La comision de la nota no puede ser mayor a la comision de la factura", vbCritical
                Exit Sub
            End If
        End If
        
        serieFac = Trim(rsBD!bon_fac_serie)  'poner donde si guarde
       
        If rsBD!pedido = 0 Then
        
            'Checar si es prefactura
            sqls = " select * from prefacturas" & _
                   " Where bodegarefer = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                   " And refer = " & Val(txtFactura.Text) & " And Cliente = " & Val(txtCliente.Text)
                   
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If rsBD.EOF Then  'Es prefactura
            
                sqls = "Exec Sp_ChecaNotasCreditoOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & _
                    Val(txtCliente.Text) & "," & Val(txtFactura.Text) & ""
                Set rsBD = New ADODB.Recordset
                rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
                
                If Not rsBD.EOF Then
                    MsgBox "Ya existe una nota de credito para esa factura, Favor de verificar", vbInformation
                    Call LimpiarControles(Me)
                    rsBD.Close
                    Set rsBD = Nothing
                    Exit Sub
                Else
                    rsBD.Close
                    Set rsBD = Nothing
                End If
            End If
        End If
    End If
    
    If chkOI.value = 1 Then 'si esta seleccionao otros ingresos
            sqls = "Exec Sp_ChecaNotasCreditoOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & _
                Val(txtCliente.Text) & "," & Val(txtFactura.Text) & ""
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
                MsgBox "Ya existe una nota de credito para esa factura, Favor de verificar", vbInformation
                Call LimpiarControles(Me)
                rsBD.Close
                Set rsBD = Nothing
                Exit Sub
            Else
                rsBD.Close
                Set rsBD = Nothing
            End If
    End If

    If optConc(0).value = False And optConc(1).value = False And optConc(2).value = False And optConc(3).value = False Then
        MsgBox "No ha seleccionado ningun concepto, no se puede generar la nota de crédito"
        Exit Sub
    End If
    
    '******
    If chkOI.value = 1 And optConc(2).value = True And Product >= 1 Then 'And optProd(2).Value = True Then
       sql = "SELECT ISNULL(TipoBon,6) TipoBon,Serie  FROM fm_clientes_movimientos"
       sql = sql & " Where BODEGA=" & cboBodegas.ItemData(cboBodegas.ListIndex)
       sql = sql & " AND Refer_Apl=" & Val(txtFactura)
       sql = sql & " AND Cliente=" & Val(txtCliente)
       Set rsBD = New ADODB.Recordset
       rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
       If Not rsBD.EOF Then
          tpb = rsBD!TipoBon
          serieFac = rsBD!serie
       End If
    End If
    '******
    
    sqls = " select * from folios"
    sqls = sqls & " where bodega   = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    sqls = sqls & " and tipo = 'NC'"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        serie = Trim(rsBD!Prefijo)
        txtFolio = Val(rsBD!consecutivo) + 1
        
        If optConc(0).value = True Then
            Concepto = 0
            TIPO_MOV = 79
        ElseIf optConc(1).value = True Then
            Concepto = 1
            TIPO_MOV = 81
        ElseIf optConc(2).value = True Then
            Concepto = 2
            TIPO_MOV = 79
        ElseIf optConc(3).value = True Then
            Concepto = 3
            TIPO_MOV = 80
        End If
        
        If chkOI.value = 1 Then
            Concepto = 2
            TIPO_MOV = 89
        End If
        
'        If optProd(0).Value = True Then
'            TipoBon = 1
'        ElseIf optProd(1).Value = True Then
'            TipoBon = 2
'        ElseIf optProd(2).Value = True Then
'            TipoBon = tpb
'        Else
'            TipoBon = 0
'        End If
         TipoBon = Product
        
        'If optProd(0).Value = True Or optProd(2).Value = True Then
        If Product = 1 Then
            importe = Val(txtTotalExe)
            Comision = Val(txtComExe)
        Else
            importe = Val(txtTotalGra)
            Comision = Val(txtComGra)
        End If
        
        sqls = " EXEC sp_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & ", @Cliente      = " & txtCliente
        sqls = sqls & vbCr & ", @Fecha        =    '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = " & TIPO_MOV
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(txtserie) & "'"  'Serie Nota Credito
        sqls = sqls & vbCr & ", @Refer        = " & txtFolio
        sqls = sqls & vbCr & ", @Refer_Apl    = " & txtFactura
        sqls = sqls & vbCr & ", @CarAbo       = 'A'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Importe      = " & importe
        '@Fecha_vento
        '@Vendedor
        sqls = sqls & vbCr & ", @IVA= " & Val(txtIvaComGra)
        '@IEPS
        sqls = sqls & vbCr & ", @Cuenta_origen = " & txtFolio
        sqls = sqls & vbCr & ", @CreditoFac = 'N'"
        '@Ficha
        sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = " & Usuario
        sqls = sqls & vbCr & ", @TipoBon = " & TipoBon
        cnxBD.Execute sqls, intRegistros
        
        'If chkOI.Value = 1 And optConc(2).Value = True And optProd(2).Value = True Then
        If chkOI.value = 1 And optConc(2).value = True Then
        sqls = " EXEC sp_fm_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & ", @Cliente      = " & Val(txtCliente)
        sqls = sqls & vbCr & ", @Fecha        = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = " & TIPO_MOV
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(txtserie) & "'"
        sqls = sqls & vbCr & ", @Refer        = 2"
        sqls = sqls & vbCr & ", @Refer_Apl    = " & txtFactura
        sqls = sqls & vbCr & ", @CarAbo       = 'A'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Importe      = " & importe
        sqls = sqls & vbCr & ", @iva      = " & Val(txtIvaComGra)
        sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Fecha_Mov =  '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = '" & gstrUsuario & "'"
        sqls = sqls & vbCr & ", @TipoBon = " & TipoBon
        cnxBD.Execute sqls, intRegistros
        End If
        
        '---Verificar ke no este vacio el campo seriefac
        If serieFac = "" Then
           serif = "select BON_FAC_SERIE,bon_Fac_pedido from bon_factura where bon_fac_sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex) & "  And bon_fac_numero = " & txtFactura
            Set rsBD = New ADODB.Recordset
            rsBD.Open serif, cnxBD, adOpenForwardOnly, adLockReadOnly
            If Not rsBD.EOF Then
                If rsBD!bon_fac_pedido <> 0 Then
                   serieFac = Trim(rsBD!bon_fac_serie)
                End If
            End If
        End If
        '-------------------------------------------
        sqls = "exec sp_Notas "
        sqls = sqls & vbCr & "   @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & "  ,@Folio        = " & txtFolio
        sqls = sqls & vbCr & "  ,@Serie        = '" & Trim(serie) & "'"
        sqls = sqls & vbCr & "  ,@cliente      = " & txtCliente & ""
        sqls = sqls & vbCr & "  ,@Factura      = " & txtFactura
        sqls = sqls & vbCr & "  ,@TipoMov      = " & TIPO_MOV
        sqls = sqls & vbCr & "  ,@Concepto      = " & Concepto
        sqls = sqls & vbCr & "  ,@FechaEmi      = '" & txtFecha & "'"
        'If optProd(2).Value = True Then
        If Product >= 1 Then
           sqls = sqls & vbCr & "  ,@TpoBon      = " & tpb & ""
        Else
           sqls = sqls & vbCr & "  ,@TpoBon      = " & TipoBon & ""
        End If
        
       'If optProd(0).Value = True Or optProd(2).Value = True Then
       If Product = 1 Then
          sqls = sqls & vbCr & "  ,@Bonexe      = " & Val(txtExe)
          sqls = sqls & vbCr & "  ,@Bongra     = 0"
          sqls = sqls & vbCr & "  ,@Ivagra      = 0"
          sqls = sqls & vbCr & "  ,@Comision      = " & Val(txtComExe)
          sqls = sqls & vbCr & "  ,@IvaCom      = " & Val(txtIvaComExe)
          sqls = sqls & vbCr & "  ,@Valor      = " & txtTotalExe
       ElseIf Product >= 2 Then
          sqls = sqls & vbCr & "  ,@Bonexe      = 0"
          sqls = sqls & vbCr & "  ,@Bongra     = " & Val(txtGra)
          sqls = sqls & vbCr & "  ,@Ivagra      = " & Val(txtIvaGra)
          sqls = sqls & vbCr & "  ,@Comision      = " & Val(txtComGra)
          sqls = sqls & vbCr & "  ,@IvaCom      = " & Val(txtIvaComGra)
          sqls = sqls & vbCr & "  ,@Valor      = " & Val(txtTotalGra)
       End If
       
                       
        sqls = sqls & vbCr & "  ,@status      = 0"
        sqls = sqls & vbCr & "  ,@Reembolso      = " & txtRem
        sqls = sqls & vbCr & "  ,@TipoBon      = " & tpb & ""
        sqls = sqls & vbCr & "  ,@SerieFac      = '" & serieFac & "'"
        sqls = sqls & vbCr & "  ,@Usuario      = '" & Usuario & "'"
        sqls = sqls & vbCr & "  ,@Plaza      = '" & txtPlaza & " '"
                
        cnxBD.Execute sqls, intRegistros
        
        sqls = "update folios set consecutivo = " & Val(txtFolio) & _
               " where tipo = 'NC'"
               
        sqls = sqls & " and bodega  = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        
        cnxBD.Execute sqls, intRegistros
        
    Else
        MsgBox "Error al buscar el folio, verifique con Sistemas"
        Exit Sub
    End If
    
    rsBD.Close
    Set rsBD = Nothing
    Observ = UCase(Trim(txtobs.Text))
    Call doGenArchFE_NCA(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(serie), Val(txtFolio.Text), Val(txtFolio.Text), ncNormal_OI)
     
    MsgBox "Nota de credito " & txtFolio & " generada!"
    NumeroNota = Val(txtFolio) + 1
    
    LimpiarControles Me
    
    lblNombre.Caption = ""
    txtFecha = Format(Date, "MM/DD/YYYY")
    
    frProd.Enabled = True 'False
    frConcepto.Enabled = False
    frGrav.Enabled = False
    frExentos.Enabled = False
    frBotones.Enabled = False
    Frame3.Enabled = False
    optLectura(0).value = True
    
    For i = 0 To 3
        optConc(i).Enabled = True
    Next i
    
    If TipoAcceso = "Factura" Then
        CargaBodegas cboBodegas
    Else
        CargaBodegasServ cboBodegas
    End If
    BuscaSigFolio (cboBodegas.ItemData(cboBodegas.ListIndex))
    Frame1.Enabled = True
    txtRem.SetFocus
    
    Call CargaComboBE(cboProductos, "sp_sel_productobe 'BE','','Todos'")
    Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & Trim(cboProductos.Text) & " ','Leer'")
    cboProductos.Text = UCase("Winko Mart")
    Exit Sub

err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, "ERROR", Usuario, "frmNotas.cmdGrabar")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Notas de Credito"
   Resume Next
   MsgBar "", False
End Sub
Sub ImprimeNota(Bodega As Integer, Folio As Integer)
    
    sqls = "select a.folio, a.serie, a.cliente, b.nombre, b.RFC, b.domicilio,b.colonia, b.codigopostal,"
    sqls = sqls & " c.descripcion AS POBLA ,a.factura, a.valor, a.FechaEmi,"
    sqls = sqls & " a.BonExe , a.BonGra, a.IvaGra, a.Comision, a.IvaCom, a.plaza, d.desccorta Estado"
    sqls = sqls & " from   notascre a, clientes b, poblaciones c, estados d"
    sqls = sqls & " Where a.Cliente = b.Cliente"
    sqls = sqls & " and b.poblacion = c.poblacion"
    sqls = sqls & " and a.bodega = " & Bodega
    sqls = sqls & " and a.folio = " & Folio
    sqls = sqls & " and c.estado  = d.estado"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        
        sql = " select fecha from clientes_movimientos"
        sql = sql & " Where Bodega = " & Bodega
        sql = sql & " and refer_Apl = " & rsBD!Factura
        sql = sql & " and cliente = " & rsBD!cliente
        sql = sql & " and tipo_mov = 10"
    
        Set rsBD2 = New ADODB.Recordset
        rsBD2.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    
        If Not rsBD2.EOF Then
            FechaFac = rsBD2!Fecha
        Else
            Fecha = ""
        End If
    
        If NumeroNota = 1 Then
           Print #2,
        End If
        
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        
        Print #2, Tab(29); "NOTA DE CREDITO     FOLIO: " & Format(rsBD!serie, "!@@") & " " & Format(rsBD!Folio, "!@@@@@@")
        Print #2,
        Print #2, Tab(17); "CIA.: " & Format(rsBD!cliente, "@@@@@@")
        Print #2, Tab(17); "SR.(ES): " & RTrim(Mid(rsBD!Nombre, 1, 30)) + LTrim(Mid(rsBD!Nombre, 31, 80))
        Print #2, Tab(26); Trim(rsBD!Plaza)
        Print #2, Tab(17); "R.F.C.: " & Trim(rsBD!Rfc); Tab(51); "VALOR: $ " & Format(Format(rsBD!valor, "###,###,###,###.00"), "@@@@@@@@@@@@@@@")
        Print #2, Tab(17); Left(rsBD!Domicilio, 33); Tab(51); "FECHA: " & Format(rsBD!fechaemi, "YYYY-MM-DD")
        Print #2, Tab(17); Left(rsBD!Colonia, 30)
        Print #2, Tab(17); Trim(Left(rsBD!POBLA, 40)) & "," & Trim(rsBD!estado); Tab(51); "C.P.:" & rsBD!CodigoPostal
        Print #2, Tab(17); "CON ESTA FECHA HEMOS ABONADO EN SU CUENTA LO SIGUIENTE:"
        Print #2, Tab(17); "| ----------------------------------------------------------------- |"
        Print #2, Tab(17); "|            DISTRIBUCION          |             CONCEPTO           |"
        Print #2, Tab(17); "| ----------------------------------------------------------------- |"
        
         If optConc(0).value = True Then
            Print #2, Tab(17); "| BONOS EXENTOS:  $" & Format(Format(IIf(IsNull(rsBD!BONEXE), 0, rsBD!BONEXE), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| BONOS DEVUELTOS SEGUN DETALLE  |"
            Print #2, Tab(17); "| BONOS GRAVADOS: $" & Format(Format(IIf(IsNull(rsBD!BONGRA), 0, rsBD!BONGRA), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| ANEXO                          |"
            
        ElseIf optConc(1).value = True Then
            Print #2, Tab(17); "| BONOS EXENTOS:  $" & Format(Format(IIf(IsNull(rsBD!BONEXE), 0, rsBD!BONEXE), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| BONIFICACION ESPECIAL POR      |"
            Print #2, Tab(17); "| BONOS GRAVADOS: $" & Format(Format(IIf(IsNull(rsBD!BONGRA), 0, rsBD!BONGRA), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| COMISION MAL APLICADA          |"
            
        ElseIf optConc(2).value = True Then
            Print #2, Tab(17); "| BONOS EXENTOS:  $" & Format(Format(IIf(IsNull(rsBD!BONEXE), 0, rsBD!BONEXE), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| CANCELACION DE FACTURA No.     |"
            Print #2, Tab(17); "| BONOS GRAVADOS: $" & Format(Format(IIf(IsNull(rsBD!BONGRA), 0, rsBD!BONGRA), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "|                                |"
            
        ElseIf optConc(3).value = True Then
            Print #2, Tab(17); "| BONOS EXENTOS:  $" & Format(Format(IIf(IsNull(rsBD!BONEXE), 0, rsBD!BONEXE), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| BONIFICACION SOBRE COMPRA      |"
            Print #2, Tab(17); "| BONOS GRAVADOS: $" & Format(Format(IIf(IsNull(rsBD!BONGRA), 0, rsBD!BONGRA), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "|                                |"
        End If
        
        sqls = " select bodega, Impuestointerior as porc from bodegas"
        sqls = sqls & " Where Bodega = " & Bodega
        
    
        Set rsiva = New ADODB.Recordset
        rsiva.Open sqls, cnxBD, adOpenStatic, adLockReadOnly
        
        If Not rsiva.EOF Then
            LEYIVA = rsiva!Porc
        Else
            LEYIVA = 0
        End If
    
        
        Print #2, Tab(17); "| IVA BONOS " & LEYIVA & "% : $" & Format(Format(IIf(IsNull(rsBD!ivagra), 0, rsBD!ivagra), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| FACTURA:" & Format(rsBD!Factura, "@@@@@@"); Tab(74); Format(FechaFac, "YYYY-MM-DD"); " |"
        Print #2, Tab(17); "| COMISION " & LEYIVA & "%  : $" & Format(Format(IIf(IsNull(rsBD!Comision), 0, rsBD!Comision), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "|                                |"
        
        
        Leyenda (rsBD!valor), ""
        letrero = "(" & letrero & ")"
        texto1 = Mid(letrero, 1, 30)
        texto2 = Mid(letrero, 31, 30)
        texto3 = Mid(letrero, 62, 30)
        
        
                
        Print #2, Tab(17); "| IVA COMIS. " & LEYIVA & "% :$" & Format(Format(rsBD!Ivacom, "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| " & Format(texto1, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " |"
        Print #2, Tab(17); "| T O T A L:      $" & Format(Format(rsBD!valor, "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| " & Format(texto2, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " |"
        Print #2, Tab(17); "|                                  "; Tab(52); "| " & Format(IIf(texto3 = "", " ", texto3), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " |"
        Print #2, Tab(17); "|                                  |                                |"
        Print #2, Tab(17); "| -------------------------------- | ------------------------------ |"
        Print #2,
          
   End If
            
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
   Reimp_FCNC = 3
   frmFact.Show 1
   'Call doGenArchFE(5, "CE", 70, 70, 3)
   
   'Call doGenArchFE_OI(5, "OE", 253, 253)
End Sub

Private Sub cboProductos_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & Trim(cboProductos.Text) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Private Sub Form_Load()
    
    Limpia_Campos
    
    CargaBodegas cboBodegas
    chkOI.value = 0
    
    sql = "sp_Bodegas_sel " & cboBodegas.ItemData(cboBodegas.ListIndex)

    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

    If Not rsBD.EOF Then
       txtFecha = Format(rsBD!FechaOp, "mm/dd/yyyy")
    Else
       MsgBox "Error al cargar la fecha de facturación, no sepueden cargar los pedidos de esta sucursal"
       Exit Sub
    End If
    
          
'   Abre el archivo para generar las notasde crédito facturada
    
    BuscaSigFolio (cboBodegas.ItemData(cboBodegas.ListIndex))
    optLectura(0).value = True
    
    Call CargaComboBE(cboProductos, "sp_sel_productobe 'BE','','Todos'")
    Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & Trim(cboProductos.Text) & " ','Leer'")
    cboProductos.Text = UCase("Winko Mart")
  
    frProd.Enabled = True 'False
    frConcepto.Enabled = False
    frGrav.Enabled = False
    frExentos.Enabled = False
    frBotones.Enabled = False
    Frame3.Enabled = False
    ncNormal_OI = 3 'NORMAL
    Call STAT
    
End Sub

Sub Limpia_Campos()

    frExentos.Visible = False
    frGrav.Visible = False

    LimpiarControles Me
End Sub
Sub BuscaSigFolio(Bodega As Integer)
If Val(Bodega) >= 1 Then
    
    sqls = " select * from folios"
    sqls = sqls & " where bodega   = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
    sqls = sqls & " and tipo = 'NC'"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rsBD.EOF Then
      txtFolio = Val(rsBD!consecutivo) + 1
    End If
End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call cmdSalir_Click
End Sub

Private Sub lstPlazas_DblClick()
    lblPlaza = lstPlazas.Text
    lstPlazas.Visible = False
End Sub

Private Sub optConc_Click(Index As Integer)

    If (Index = 2 And txtRem <> "0") Or (Index = 2) Then  'Cancelación Total  optProd(2).Value
        sql = " select * from clientes_movimientos"
        sql = sql & " Where Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sql = sql & " and refer_Apl = " & Val(txtFactura)
        sql = sql & " and serie = '" & txtserie & "'"
        sql = sql & " and tipo_mov = 10 and cliente = " & txtCliente.Text

        Set rsBD2 = New ADODB.Recordset
        rsBD2.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

        If Not rsBD2.EOF Then
            'If optProd(0).Value = True Or optProd(2).Value = True Then 'despensa o BE
            If Product = 1 Then
                If (Val(rsBD2!bon_fac_bonexe) = Val(txtExe)) Or Product = 1 Then 'optProd(2).Value
                    txtExe = rsBD2!bon_fac_bonexe
                    txtComExe = IIf(IsNull(rsBD2!bon_fac_comision), 0, rsBD2!bon_fac_comision)
                    txtIvaComExe = IIf(IsNull(rsBD2!bon_fac_ivacom), 0, rsBD2!bon_fac_ivacom)
                    txtTotalExe = Format(Val(txtExe) + Val(txtIvaComExe) + Val(txtComExe), "######.00")
                ElseIf Val(rsBD2!bon_fac_bonexe) > Val(txtExe) Then
                    resp = MsgBox("El valor de los bonos leidos es menor al valor de la factura " & Chr(13) & " No se puede hacer una cancelación total de la factura." & Chr(13) & "Desea hacer una cancelacion parcial?", vbYesNo)
                    If resp = vbYes Then
                        txtTotalExe = Format(Val(txtExe), "######.00")
                        optConc(0).value = True
                    Else
                        optConc(0).value = False
                        optConc(2).value = False
                    End If
                ElseIf Val(rsBD2!bon_fac_bonexe) < Val(txtExe) Then
                     resp = MsgBox("El valor de la factura es menor al valor de los bonos leídos " & Chr(13) & "Desea hacer la nota solo por el valor de la factura?", vbYesNo)
                     If resp = vbYes Then
                        txtComExe = rsBD2!bon_fac_comision
                        txtIvaComExe = rsBD2!bon_fac_ivacom
                        txtTotalExe = Format(Val(txtExe) + Val(txtIvaComExe) + Val(txtComExe), "######.00")
                     End If
                End If
            ElseIf Product <> 1 Then 'gas  optProd(1).Value = True
                    txtGra = Val(rsBD2!bon_fac_bongra)
                    txtIvaGra = rsBD2!bon_fac_ivagra
                    txtComGra = rsBD2!bon_fac_comision
                    txtIvaComGra = rsBD2!bon_fac_ivacom
                    txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "#######.00")
             
            
                    
                If Val(rsBD2!bon_fac_bongra) + Val(rsBD2!bon_fac_ivagra) = Val(txtGra) Then
                    txtGra = Val(rsBD2!bon_fac_bongra)
                    txtIvaGra = rsBD2!bon_fac_ivagra
                    txtComGra = rsBD2!bon_fac_comision
                    txtIvaComGra = rsBD2!bon_fac_ivacom
                    txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "#######.00")
                ElseIf Val(rsBD2!bon_fac_bongra) > Val(txtGra) And Val(txtRem.Text) Then
                    resp = MsgBox("El valor de los bonos leidos es menor al valor de la factura . No se puede hacer una cancelación total de la factura. Desea hacer una cancelacion parcial?", vbYesNo)
                    If resp = vbYes Then
                        txtTotalGra = Val(txtGra)
                        optConc(0).value = 0
                    Else
                        txtGra = 0
                        txtIvaGra = 0
                        txtComGra = 0
                        txtIvaComGra = 0
                        txtTotalGra = 0
                    End If
                ElseIf Val(rsBD2!bon_fac_bongra) + Val((rsBD2!bon_fac_ivagra)) < Val(txtGra) Then
                     resp = MsgBox("El valor de la factura es menor al valor de los bonos leídos " & Chr(13) & "Desea hacer la nota solo por el valor de la factura?", vbYesNo)
                     If resp = vbYes Then
                        txtGra = rsBD2!bon_fac_bongra
                        txtIvaGra = rsBD2!bon_fac_ivagra
                        txtComGra = rsBD2!bon_fac_comision
                        txtIvaComGra = rsBD2!bon_fac_ivacom
                        txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "#######.00")
                     End If
                End If
           End If
        End If
        
    ElseIf Index = 1 Or Index = 3 Then
    
        If Product = 1 Then 'optProd(0).Value = True Or optProd(2).Value = True
           
            frExentos.Enabled = True
            txtExe = 0
            txtExe.Enabled = False
            txtComExe = 0
            txtComExe.Enabled = True
            txtIvaComExe = 0
            txtIvaComExe.Enabled = False
            txtTotalExe = Format(Val(txtExe) + Val(txtIvaComExe) + Val(txtComExe), "#######.00")
            txtTotalExe.Enabled = False
            txtIvaGra.Enabled = True
            txtComGra.Enabled = True
            txtIvaComGra.Enabled = True
 
        Else
            frGrav.Enabled = True
            txtGra = 0
            txtIvaGra = 0
            txtComGra = 0
            txtIvaComGra = 0
            txtIvaGra.Enabled = True
            txtComGra.Enabled = True
            txtIvaComGra.Enabled = True

            txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "#######.00")

        End If
        
    End If
End Sub

Private Sub optLectura_Click(Index As Integer)
    txtRem.Text = ""
  '  Call LimpiafrCte
    Call LimpiafrGrav
    Call LimpiafrExe
End Sub


Private Sub optProd_Click(Index As Integer)
    If Product = 1 Then
        frExentos.Visible = True
        frGrav.Visible = False
    Else
        frExentos.Visible = False
        frGrav.Visible = True
    End If
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then

        If chkOI.value = 1 Then
           sql = "select nombre from clientesOI"
        Else
           sql = "select nombre from clientes"
        End If
        
        'SQL = "select nombre from clientes"
        sql = sql & " where cliente = " & Val(txtCliente) & ""
        sql = sql & " and bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        Set rsBD = New ADODB.Recordset
        rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

        If Not rsBD.EOF Then
            lblNombre.Caption = rsBD!Nombre
            'SendKeys "{TAB}"
            txtFactura.SetFocus
        Else
            MsgBox "Cliente no existe!, favor de verificarlo", vbCritical, "Notas de Crédito"
            txtCliente = ""
        End If
    End If
        
End Sub

Private Sub txtComExe_Change()
   If cboBodegas.Text <> "" Then
   
        sqls = " select bodega, Impuestointerior as porc from bodegas"
        sqls = sqls & " Where Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
          
        Set rsiva = New ADODB.Recordset
        rsiva.Open sqls, cnxBD, adOpenStatic, adLockReadOnly
        
        If Not rsiva.EOF Then
            txtIvaComExe = Format(Val(txtComExe) * (rsiva!Porc / 100), "#####0.00")
        Else
            txtIvaComExe = 0
            MsgBox "Error al calcular el iva", vbCritical
        
        End If
    
        txtTotalExe = Format(Val(txtExe) + Val(txtIvaComExe) + Val(txtComExe), "######0.00")
    End If
End Sub

Private Sub txtComExe_KeyPress(KeyAscii As Integer)
    If ValidaNumericos(KeyAscii, txtExe, 3) Then entertab (KeyAscii)
End Sub

Private Sub txtComGra_Change()
    
    If cboBodegas.Text <> "" Then
        sqls = " select bodega, Impuestointerior as porc from bodegas"
        sqls = sqls & " Where Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
            
        
        Set rsiva = New ADODB.Recordset
        rsiva.Open sqls, cnxBD, adOpenStatic, adLockReadOnly
        
        If Not rsiva.EOF Then
            txtIvaComGra = Val(txtComGra) * (rsiva!Porc / 100)
        Else
            txtIvaComGra = 0
            MsgBox "Error al calcular el iva", vbCritical
        
        End If
    
        txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "######0.00")
    End If
End Sub

Private Sub txtExe_Change()
    txtTotalExe = Format((Val(txtExe) + Val(txtIvaComExe) + Val(txtComExe)), "######0.00")
End Sub

Private Sub txtExe_KeyPress(KeyAscii As Integer)

If ValidaNumericos(KeyAscii, txtExe, 3) Then entertab (KeyAscii)

End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
        KeyAscii = 0

    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chkOI.value = 0 Then 'si no es de otros ingresos cheka en bon_factura
            sql = " select bon_fac_bonexe, bon_fac_bongra, bon_fac_ivagra, bon_fac_comision, "
            sql = sql & " bon_fac_ivacomis as ivacomision, bon_fac_tpobon, bon_fac_pedido , bon_Fac_status from bon_factura"
            sql = sql & " Where bon_fac_sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex)
            sql = sql & " and bon_fac_numero  = " & Val(txtFactura)
            sql = sql & " and bon_fac_serie = '" & Trim(txtserie) & "'"
            sql = sql & " and bon_fac_cliente  = " & Val(txtCliente) & ""

            Set rsBD = New ADODB.Recordset
            rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If rsBD.EOF Then
                MsgBox "La factura no existe o no corresponde al cliente", vbCritical
                txtFactura.Text = ""
                Exit Sub
            Else
                If rsBD!bon_Fac_Status = 2 Then
                    MsgBox "La Factura esta cancelada, no se puede hacer una nota de crédito", vbCritical
                    txtFactura.Text = ""
                    Exit Sub
                End If
            End If
        Else  'si es de otros ingresos busca en fm_facturas
            sql = "SELECT * FROM fm_facturas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & " "
            sql = sql & " And Cliente = " & Val(txtCliente.Text) & " and factura = " & Val(txtFactura) & " "
            
            Set rsBD = New ADODB.Recordset
            rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
                                
            If rsBD.EOF Then
                MsgBox "La factura no existe o no corresponde al cliente", vbCritical
                txtFactura.Text = ""
                Exit Sub
            Else
                If rsBD!Status = 2 Then
                    MsgBox "La Factura esta cancelada, no se puede hacer una nota de crédito", vbCritical
                    txtFactura.Text = ""
                    Exit Sub
                End If
                
                Frame2.Visible = True
                txtRem.Text = 0
                txtserie.Text = rsBD!serie
                txtComExe.Text = 0
                txtExe.Text = Val(rsBD!Subtotal)
                txtIvaComExe.Text = Val(rsBD!iva)
                txtTotalExe.Text = rsBD!Subtotal + rsBD!iva
                optProd(2).value = True
                frConcepto.Enabled = True
                optConc(0).Enabled = False
                optConc(1).Enabled = False
                optConc(3).Enabled = False
                cmdAjuste.Enabled = False
                Frame1.Enabled = False
                frBotones.Enabled = True
                Frame3.Enabled = True
            End If
        End If
    End If
        
End Sub

Private Sub txtFactura_LostFocus()
    txtFactura_KeyPress (13)
End Sub
                                                                                                                                                                                                                                                            
Private Sub txtGra_Change()
    txtTotalGra = Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra)
End Sub

Private Sub txtGra_KeyPress(KeyAscii As Integer)
   If ValidaNumericos(KeyAscii, txtGra, 3) Then entertab (KeyAscii)

End Sub

Private Sub txtIvaComExe_Change()
   txtTotalExe = Format(Val(txtExe) + Val(txtIvaComExe) + Val(txtComExe), "######0.00")
End Sub

Private Sub txtIvaComExe_KeyPress(KeyAscii As Integer)
    If ValidaNumericos(KeyAscii, txtExe, 3) Then entertab (KeyAscii)
End Sub

Private Sub txtIvaComGra_Change()
    txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "######0.00")
End Sub

Private Sub txtIvaGra_Change()
    txtTotalGra = Format(Val(txtGra) + Val(txtIvaGra) + Val(txtComGra) + Val(txtIvaComGra), "######0.00")
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPass = "Sup" Then
            frExentos.Enabled = True
        Else
            MsgBox "Password incorrecto!, favor de verificarlo", vbCritical
        End If
        
        frPass.Visible = False
    End If
            
End Sub

Private Sub txtRem_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtTotalExe_KeyPress(KeyAscii As Integer)
    If ValidaNumericos(KeyAscii, txtExe, 3) Then entertab (KeyAscii)
End Sub

Sub LimpiafrCte()
    Call CboPosiciona(cboBodegas, gnBodega)
    txtCliente.Text = ""
    lblNombre.Caption = ""
    txtFactura.Text = ""
End Sub

Sub LimpiafrGrav()
    txtGra.Text = ""
    txtIvaGra.Text = ""
    txtComGra.Text = ""
    txtIvaComGra.Text = ""
    txtTotalGra.Text = ""
    frGrav.Enabled = False
End Sub

Sub LimpiafrExe()
    txtExe.Text = ""
    txtComExe.Text = ""
    txtIvaComExe.Text = ""
    txtTotalExe.Text = ""
    frExentos.Enabled = False
End Sub

Sub STAT()
    sql = "SELECT * FROM CONFIGBODEGAS WHERE Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       If Month(Date) <> Month(rsBD!FechaCierre) Then
          cmdGrabar.Enabled = False
       Else
          cmdGrabar.Enabled = True
       End If
    Else
       MsgBox "Error no hay fecha de cierre actual para esa Sucursal", vbCritical, "Sin fecha"
    End If
End Sub



