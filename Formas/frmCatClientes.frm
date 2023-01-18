VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCatClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de Clientes"
   ClientHeight    =   6045
   ClientLeft      =   -1785
   ClientTop       =   1470
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   8895
      Begin VB.TextBox txtPB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   51
         Tag             =   "Det"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtedo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   50
         Tag             =   "Det"
         Top             =   3360
         Width           =   3780
      End
      Begin VB.TextBox txtClienteDisp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   47
         Tag             =   "Det"
         Top             =   2040
         Width           =   1515
      End
      Begin VB.CommandButton cmdClienteFac 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         Picture         =   "frmCatClientes.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox txtClienteFac 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   44
         Tag             =   "Det"
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txtMail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   42
         Tag             =   "Det"
         Top             =   1320
         Width           =   2595
      End
      Begin VB.Frame Frame3 
         Caption         =   "Costos Tarjetas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5280
         TabIndex        =   39
         Top             =   3000
         Width           =   2775
         Begin FPSpread.vaSpread spdcostos 
            Height          =   735
            Left            =   240
            OleObjectBlob   =   "frmCatClientes.frx":1DFC
            TabIndex        =   40
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox txtRFC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         Tag             =   "Det"
         Top             =   600
         Width           =   2595
      End
      Begin VB.TextBox txtCodigoPostal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   26
         Tag             =   "Det"
         Top             =   240
         Width           =   870
      End
      Begin VB.TextBox txtTelefono 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   25
         Tag             =   "Det"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPoblacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Tag             =   "Det"
         Top             =   2445
         Width           =   585
      End
      Begin VB.ComboBox cboEstados 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1935
         Style           =   1  'Simple Combo
         TabIndex        =   23
         Text            =   "cboEstados"
         Top             =   2895
         Width           =   2970
      End
      Begin VB.TextBox txtEntreCalles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   22
         Tag             =   "Det"
         Top             =   2040
         Width           =   3690
      End
      Begin VB.TextBox txtColonia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   21
         Tag             =   "Det"
         Top             =   1560
         Width           =   3690
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   20
         Tag             =   "Det"
         Top             =   240
         Width           =   3660
      End
      Begin VB.TextBox txtDomicilio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   19
         Tag             =   "Det"
         Top             =   1080
         Width           =   3660
      End
      Begin VB.ComboBox cboVendedores 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCatClientes.frx":2076
         Left            =   6720
         List            =   "frmCatClientes.frx":2078
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "Det"
         Top             =   960
         Width           =   1860
      End
      Begin VB.TextBox txtEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2880
         Width           =   585
      End
      Begin VB.TextBox txtVendedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   16
         Tag             =   "Det"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtNombre1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         Tag             =   "Det"
         Top             =   600
         Width           =   3675
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dias de espera para dispersion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5160
         TabIndex        =   52
         Top             =   2700
         Width           =   2625
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Nom. Edo. Cuenta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         TabIndex        =   49
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente Dispersar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5040
         TabIndex        =   48
         Top             =   2115
         Width           =   1485
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente a Facturar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   5040
         TabIndex        =   45
         Top             =   1755
         Width           =   1725
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5040
         TabIndex        =   43
         Top             =   1400
         Width           =   540
      End
      Begin VB.Label lblRFC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5040
         TabIndex        =   38
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblTelefono 
         Appearance      =   0  'Flat
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5040
         TabIndex        =   37
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblCP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "C. P.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7200
         TabIndex        =   36
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblPoblacion 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   35
         Tag             =   "Det"
         Top             =   2445
         Width           =   2955
      End
      Begin VB.Label lblEstado 
         Appearance      =   0  'Flat
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   105
         TabIndex        =   34
         Top             =   2910
         Width           =   1200
      End
      Begin VB.Label lblEntreCalles 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Entre calles:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   33
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label lblColonia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   32
         Top             =   1515
         Width           =   780
      End
      Begin VB.Label lblPobl 
         Appearance      =   0  'Flat
         Caption         =   "Población:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   105
         TabIndex        =   31
         Top             =   2475
         Width           =   1020
      End
      Begin VB.Label lblDomicilio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   30
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   100
         TabIndex        =   29
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblVendedor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5040
         TabIndex        =   28
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdEntrega 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2520
         Picture         =   "frmCatClientes.frx":207A
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Captura de datos de entrega"
         Top             =   220
         Width           =   400
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3120
         Picture         =   "frmCatClientes.frx":24BC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   220
         Width           =   400
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         Picture         =   "frmCatClientes.frx":25BE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   220
         Width           =   400
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1320
         Picture         =   "frmCatClientes.frx":26C0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   220
         Width           =   400
      End
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Picture         =   "frmCatClientes.frx":27C2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   220
         Width           =   400
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&X"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   220
         Width           =   400
      End
   End
   Begin VB.Frame fraPrincipal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8895
      Begin VB.ComboBox cboSucursal 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Key"
         Top             =   360
         Width           =   2280
      End
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Tag             =   "Key"
         Top             =   360
         Width           =   780
      End
      Begin MSMask.MaskEdBox mskFechaBanco 
         Height          =   345
         Left            =   5040
         TabIndex        =   41
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5040
         TabIndex        =   13
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label lblFecIngresoV 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblFecIngresoT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3315
         TabIndex        =   5
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label lblBodega 
         Alignment       =   2  'Center
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmCatClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim blnNew As Boolean
Dim CanalVta As Integer, TipoVta As Integer

'Variables de Datos de Créditos
Dim BodegaAutPolCre As Boolean

'Variables de Datos de Distribucion
Dim nTolera As Integer

Private mstrTipo As String
Private mlngCliente As Long

Public Property Let tipo(sTipo As String)
   mstrTipo = sTipo
End Property

Public Property Get tipo() As String
   tipo = mstrTipo
End Property
Public Property Let cliente(NCliente As Long)
   mlngCliente = NCliente
End Property
Public Property Get cliente() As Long
   cliente = mlngCliente
End Property


Sub limpiacajas()
Dim i As Byte
Dim ctl As Object
For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
       ctl.Text = ""
    End If
Next
lblFecIngresoV.Caption = ""
mskFechaBanco.Text = "__/__/____"
For i = 0 To 2
    spdcostos.Col = 1
    spdcostos.Row = i
    spdcostos.Text = ""
    spdcostos.Col = 2
    spdcostos.Row = i
    spdcostos.Text = ""
Next
End Sub

Sub InicializaForma()
   'CargaDatos
   '--------------------
   'txtCliente.Text = ""
   Call limpiacajas
   '--------------------
   txtCliente.Enabled = True
   'txtCliente.Locked = False
   MsgBar "Cargando catálogos", True
   Call CentraForma(Me)
   TipoBusqueda = "ClienteBE"
   
   Call CargaBodegas(CboSucursal)
   Call CargaEstados(cboEstados)
   Call CargaVendedores(cboVendedores)
    
   cmdActualizar.Enabled = False
   cmdAbrir.Enabled = False
'   tblComisiones.Enabled = False
   CboSucursal.Enabled = False
   MsgBar "", False
End Sub


Private Sub cmdClienteFac_Click()
Dim frmConsulta As New frmBusca_Cliente
    
    frmConsulta.Bodega = CboSucursal.ItemData(CboSucursal.ListIndex)
    frmConsulta.Show vbModal
   
    If frmConsulta.cliente > 0 Then
       txtClienteFac = frmConsulta.cliente
    End If
    Set frmConsulta = Nothing

End Sub

Private Sub cmdEntrega_Click()
  GEstado = ""
  GCd = ""
  GCliente = txtCliente
  GNombre = Trim(txtNombre) & " " & Trim(txtNombre1)
  GCalle = txtDomicilio
  GColonia = txtColonia
  GTel = txtTelefono
  GCP = txtCodigoPostal
  GEstado = txtEstado
  GCd = txtPoblacion
  frmEntrega.Show vbModal
End Sub

Private Sub Form_Load()
Dim strsql As String
Dim sPuesto As String, blnCreditos As Boolean
   Set mclsAniform = New clsAnimated
   MsgBar "Cargando catálogos", True
   
   Call CentraForma(Me)
   TipoBusqueda = "ClienteBE"
   
   Call CargaBodegas(CboSucursal)
   Call CargaEstados(cboEstados)
   Call CargaVendedores(cboVendedores)
    
   cmdActualizar.Enabled = False
   cmdAbrir.Enabled = False
   cmdEntrega.Enabled = False
'   tblComisiones.Enabled = False

   CboSucursal.Enabled = False

   MsgBar "", False
     
End Sub
Function CargaDatos() As Boolean
Dim edo As String
   CargaDatos = False
   CanalVta = 0

'   Datos del Cliente
   sqls = " "
   sqls = "EXEC   sp_clientes_sel "
   sqls = sqls & "   @Bodega  =   " & CboSucursal.ItemData(CboSucursal.ListIndex)
   sqls = sqls & " , @Cliente = " & Val(txtCliente)
    
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
   If rsBD.EOF = False Then
      lblFecIngresoV.Caption = UCase(Format(rsBD!Fecha_Alta, "DD/MMM/YYYY"))
      
      If rsBD!Nombre = Null Or IsNull(rsBD!Nombre) Then
      Else
         txtNombre = Trim(rsBD!Nombre & " " & rsBD!ApPat & " " & rsBD!ApMat)
         If Len(txtNombre) > 30 Then
            txtNombre = Mid(Trim(rsBD!Nombre & " " & rsBD!ApPat & " " & rsBD!ApMat), 1, 30)
            txtNombre1 = Mid(Trim(rsBD!Nombre & " " & rsBD!ApPat & " " & rsBD!ApMat), 31, 30)
         End If
         
         edo = Trim(txtNombre & txtNombre1) 'rsBD!nombre
         edo = Trim(edo)
        ' txtNombre1 = rsBD!nombre1
      End If
      
      If rsBD!Domicilio <> "" Then txtDomicilio = rsBD!Domicilio
      If rsBD!EntreCalles <> "" Then txtEntreCalles = rsBD!EntreCalles
      If rsBD!Colonia <> "" Then txtColonia = rsBD!Colonia
      If rsBD!Telefono <> "" Then txtTelefono = rsBD!Telefono
      If rsBD!Rfc <> "" Then txtRFC = rsBD!Rfc
      If rsBD!CodigoPostal <> "" Then txtCodigoPostal = rsBD!CodigoPostal
      If rsBD!poblacion <> 0 Then txtPoblacion = rsBD!poblacion
     ' If rsBD!NomPoblacion <> "" Then lblPoblacion = rsBD!NomPoblacion
      
      If IsNull(rsBD!estado) = False Then Call CboPosiciona(cboEstados, rsBD!estado)
      If IsNull(rsBD!vendedor) = False Then Call CboPosiciona(cboVendedores, rsBD!vendedor)
      If IsNull(rsBD!cliente_externo) = False Then txtClienteFac = rsBD!cliente_externo
      txtClienteDisp = rsBD!cliente
      txtMail.Text = IIf(IsNull(rsBD!identificacion), "", Trim(rsBD!identificacion))

      txtVendedor.Text = rsBD!vendedor
      If IsNull(rsBD!Bodega) = False Then Call CboPosiciona(CboSucursal, rsBD!Bodega)
      '------
         sqls = "SELECT ISNULL(NOMBRE,'') NOMBRE FROM CLIENTES WHERE CLIENTE=" & Val(txtCliente)
         Set rsBD = New ADODB.Recordset
         rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
         If Not rsBD.EOF Then
            If rsBD!Nombre = "" Then
               txtedo.Text = Mid(edo, 1, 34)
            Else
               txtedo.Text = Trim(Mid(rsBD!Nombre, 1, 34))
            End If
         Else
               txtedo.Text = Trim(Mid(edo, 1, 34))
         End If
       '-----
       If txtVendedor <> "" Then Call CboPosiciona(cboVendedores, Val(txtVendedor.Text))
            
   Else
      If sUbicacion = "01" Then
         If mstrTipo = "A" Then
            If MsgBox("El Cliente # " & Trim(txtCliente) & " No Existe En: " + vbCr + vbCr + "                 << Desea darlo de Alta >>", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
               CargaDatos = True
               blnNew = True
            End If
         End If
      Else
         MsgBox "El Cliente no existe,  favor de solicitar la alta al departamento de creditos!!!", vbCritical, Me.Caption
      End If
      Exit Function
   End If
       
   '---------------------------------------------
    sqls = " select clientedisp" & _
           " from clientesbe" & _
           " where clientefact = " & Val(txtCliente)
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        txtClienteDisp = rsBD!clientedisp
    End If
           
    '--------------------------------------------
    sqls = " exec sp_costostar @Bodega = 1, @Cliente = " & Val(txtCliente) & ", @Accion = 'Consultar'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    With spdcostos
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If rsBD!tipo = "B" Then
            .Text = "Banda"
        Else
            .Text = "Chip"
        End If
        .Col = 2
        .Text = CDbl(rsBD!costo)
        rsBD.MoveNext
    Next i
    End With
       
   CargaDatos = True
   
End Function


Function DatosValidos() As Boolean
   DatosValidos = True
End Function

Private Sub Inicializa()
Dim strsql As String, rstTmp As ADODB.Recordset
    
   Call Switch(Me, "Key", True)
   Call Switch(Me, "Det", False)
   Call Switch(Me, "Cargo", False)
   Call Switch(Me, "NOPOLITICA", False)
   Call Switch(Me, "Faltas", False)
   txtCliente.Text = ""
   lblFecIngresoV.Caption = ""
   
   Call CboPosiciona(cboVendedores, 0)
   txtPoblacion.Text = 977             'Monterrey
   Call CboPosiciona(cboEstados, 19)   'Nuevo Leon
   
   strsql = "SELECT Descripcion " & _
            "FROM Poblaciones " & _
            "WHERE Poblacion = " & Val(txtPoblacion.Text)
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If Not rstTmp.EOF Then
      lblPoblacion = rstTmp!descripcion
   Else
      lblPoblacion.Caption = ""
   End If
   rstTmp.Close
   Set rstTmp = Nothing
   cmdActualizar.Enabled = False
   With spdcostos
    .Col = -1
    .Row = -1
    .Action = 12
    
   End With
   
End Sub


Function ValidaLlave()
    ValidaLlave = False
    If Not IsNumeric(txtCliente) Then Exit Function
    
    ValidaLlave = True

End Function

Private Function VerificaCredito() As Boolean
 VerificaCredito = True
End Function

Private Sub cboCanales_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub cboDiasEntrega_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub cboEstados_Click()
If cboEstados.Text <> "" Then txtEstado = cboEstados.ItemData(cboEstados.ListIndex)
End Sub

Private Sub cboEstados_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub CboSucursal_Click()
Dim strsql As String
   If CboSucursal.Text <> "" Then
      strsql = "SELECT Vendedor as Clave, " & _
                        "Nombre as Descripcion " & _
               "FROM Vendedores " & _
               "WHERE Bodega = " & CboSucursal.ItemData(CboSucursal.ListIndex) & " " & _
               "AND Clasificacion = 'V' " & _
               "ORDER BY Nombre"
      Call CargaVendedores(cboVendedores)
      Call CboPosiciona(cboVendedores, 0)
   End If
    
End Sub

Private Sub cboVendedores_Click()
    If cboVendedores.Text <> "" Then txtVendedor = cboVendedores.ItemData(cboVendedores.ListIndex)
End Sub

Private Sub cboVendedores_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub chkCuentaCargo_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub chkMoroso_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub cmdAbrir_Click()
    If Trim(txtCliente.Text) <> "" Then
      If Not ValidaLlave() Then Exit Sub
      If Not CargaDatos() Then
        txtCliente.SetFocus
        Exit Sub
      End If
    Else
      blnNew = True
    End If
    Call Switch(Me, "Det", True)
    Call Switch(Me, "Key", False)
  '  chkCuentaCargo_Click
    
   cmdCancelar.Visible = True
   'cmdActualizar.Caption = IIf(blnNew, "&Agregar", "&Actualizar")
   cmdActualizar.Enabled = True
   cmdEntrega.Enabled = True
End Sub


Private Sub cmdActualizar_Click()
Dim Bodega As Integer, cliente As Integer, ClienteFac As Integer, tipo As String
Dim costo As Double

On Error GoTo ERR:
With spdcostos
    For i = 1 To .MaxRows
        .Row = i
        Bodega = CboSucursal.ItemData(CboSucursal.ListIndex)
        cliente = Val(txtCliente)
        .Col = 1
        tipo = .Text
        .Col = 2
        costo = Val(.Text)
                
        sqls = "exec sp_Costostar " & _
               " @Bodega = " & Bodega & _
               " ,@Cliente = " & cliente & _
               " ,@Tipo  = '" & Left(Trim(tipo), 1) & "'" & _
               " ,@Costo = " & costo & _
               " ,@Accion='Grabar'"

        cnxbdMty.Execute sqls, intRegistros

    Next i
End With

MsgBox "Informacion Actualizada!", vbInformation, "Clientes"


limpiacajas
txtCliente.Enabled = True
txtCliente.SetFocus
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub

Private Sub CmdBuscar_Click()
Dim frmConsulta As New frmBusca_Cliente
    
    frmConsulta.Bodega = CboSucursal.ItemData(CboSucursal.ListIndex)
    frmConsulta.Show vbModal
   
    If frmConsulta.cliente > 0 Then
       txtCliente = frmConsulta.cliente
       txtNombre = Mid(frmConsulta.Nombre, 1, 30)
       txtNombre1 = Mid(frmConsulta.Nombre, 31, 30)
       Call cmdAbrir_Click
    End If
    Set frmConsulta = Nothing
End Sub

Private Sub cmdBuscar_KeyPress(KeyAscii As Integer)
Call entertab(KeyAscii)
End Sub

Private Sub cmdCancelar_Click()
   Call InicializaForma
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub opcContado_Click()
   TipoVta = 0
   Call Switch(Me, "NOPOLITICA", False)
End Sub

Private Sub opcContado_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub opcCredito_Click()
   TipoVta = 1
   Call Switch(Me, "NOPOLITICA", True)
End Sub

Private Sub opcCredito_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub optcont_Click()
nTolera = 19
End Sub

Private Sub optFisica_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub optMoral_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub optsint_Click()
nTolera = 20
End Sub

Private Sub tblComisiones_Click()
'   tblComisiones.RowSel = tblComisiones.Row
'   tblComisiones.ColSel = tblComisiones.Col
End Sub
Private Sub txtCarteraDias_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub



Private Sub txtcliente_KeyPress(KeyAscii As Integer)
If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
    KeyAscii = 0
End If
  
  If KeyAscii = vbKeyReturn Then
    Call cmdAbrir_Click
  End If
End Sub

Private Sub txtCodigoPostal_KeyPress(KeyAscii As Integer)
If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
    KeyAscii = 0
End If

entertab KeyAscii
End Sub

Private Sub TxtColonia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
End Sub

Private Sub txtCreditoDias_KeyPress(KeyAscii As Integer)
    entertab KeyAscii
End Sub
Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
    KeyAscii = 0
End If
entertab KeyAscii
End Sub


Private Sub TxtDomicilio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
  entertab KeyAscii
End Sub

Private Sub txtEntreCalles_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
  entertab KeyAscii
End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub txtEstado_LostFocus()
If Val(txtEstado) <> 0 And IsNumeric(txtEstado) Then
     sqls = "SELECT substring(Descripcion,1,40) FROM estados"
     sqls = sqls & " WHERE estado =" & Val(txtEstado)
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
     If rsBD.EOF = False Then
       Call CboPosiciona(cboEstados, txtEstado)
     Else
       MsgBox " No existe Estado", vbOKOnly, "Validación "
       txtEstado.Clear
       txtEstado.SetFocus
     End If
Else
       txtEstado.SetFocus
End If
End Sub

Private Sub txtIdentificacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
  entertab KeyAscii
End Sub

Private Sub txtMorosidad_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
  entertab KeyAscii
End Sub
Private Sub TxtNombre1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
  entertab KeyAscii
End Sub

Private Sub txtNumCheDev_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub txtPoblacion_KeyPress(KeyAscii As Integer)
    entertab KeyAscii
End Sub

Private Sub txtPoblacion_LostFocus()
   If Val(txtPoblacion) <> 0 And IsNumeric(txtPoblacion) Then
     sqls = "SELECT Descripcion FROM Poblaciones "
     sqls = sqls & " WHERE Poblacion =" & Val(txtPoblacion)
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
     If rsBD.EOF = False Then
       lblPoblacion = rsBD!descripcion
     Else
       lblPoblacion.Caption = ""
     End If
   End If
End Sub

Private Sub TxtRFC_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub txtRutaNumero_KeyPress(KeyAscii As Integer)
entertab KeyAscii
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then 'Comillas
     KeyAscii = Asc("'")
  End If
  entertab KeyAscii
End Sub


Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
    KeyAscii = 0
End If
entertab KeyAscii
End Sub

Private Sub txtVendedor_LostFocus()
If Val(txtVendedor) <> 0 And IsNumeric(txtVendedor) Then
     sqls = "SELECT substring(NOMBRE,1,40) FROM VENDEDORES "
     sqls = sqls & " WHERE VENDEDOR =" & Val(txtVendedor) & _
     " and   Clasificacion = 'P'   "
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
     If rsBD.EOF = False Then
       Call CboPosiciona(cboVendedores, txtVendedor)
     Else
       MsgBox " No existe Vendedor", vbOKOnly, "Validación "
       txtVendedor = Empty
       txtVendedor.SetFocus
     End If
End If
End Sub
