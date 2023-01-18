VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAtencionClientes 
   Caption         =   "Atención Clientes"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmAtencionClientes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtEstado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtVendedor"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDomicilio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtColonia"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtEntreCalles"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtTelefono"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCodigoPostal"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtRFC"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMail"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label10"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label9"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label8"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label7"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label3"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblVendedor"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblDomicilio"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblPobl"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblColonia"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblEntreCalles"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblEstado"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblPoblacion"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblTelefono"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblRFC"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "Direcciones Entrega"
      TabPicture(1)   =   "frmAtencionClientes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vaSpread1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pedidos Tarjetas"
      TabPicture(2)   =   "frmAtencionClientes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vaSpread2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pedidos Dispersiones"
      TabPicture(3)   =   "frmAtencionClientes.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "vaSpread3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5295
         Left            =   -74760
         OleObjectBlob   =   "frmAtencionClientes.frx":0070
         TabIndex        =   44
         Top             =   600
         Width           =   9495
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5295
         Left            =   -74640
         OleObjectBlob   =   "frmAtencionClientes.frx":039B
         TabIndex        =   45
         Top             =   600
         Width           =   9495
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   5295
         Left            =   240
         OleObjectBlob   =   "frmAtencionClientes.frx":06F8
         TabIndex        =   46
         Top             =   480
         Width           =   9495
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   -68865
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   41
         Tag             =   "Det"
         Top             =   2220
         Width           =   3690
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   -68865
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   40
         Tag             =   "Det"
         Top             =   2700
         Width           =   3690
      End
      Begin VB.TextBox txtEstado 
         Height          =   315
         Left            =   -73785
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2580
         Width           =   3705
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -73785
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   36
         Tag             =   "Det"
         Top             =   3900
         Width           =   2595
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   -68865
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   32
         Tag             =   "Det"
         Top             =   780
         Width           =   3660
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   -68865
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   31
         Tag             =   "Det"
         Top             =   1260
         Width           =   3690
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   -68865
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   30
         Tag             =   "Det"
         Top             =   1740
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -73785
         TabIndex        =   28
         Tag             =   "Det"
         Top             =   4980
         Width           =   2595
      End
      Begin VB.TextBox txtVendedor 
         Height          =   315
         Left            =   -73785
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "Det"
         Top             =   4260
         Width           =   2610
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   315
         Left            =   -73785
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "Det"
         Top             =   780
         Width           =   3660
      End
      Begin VB.TextBox txtColonia 
         Height          =   315
         Left            =   -73785
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         Tag             =   "Det"
         Top             =   1260
         Width           =   3690
      End
      Begin VB.TextBox txtEntreCalles 
         Height          =   315
         Left            =   -73785
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         Tag             =   "Det"
         Top             =   1740
         Width           =   3690
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   -73785
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   13
         Tag             =   "Det"
         Top             =   3060
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoPostal 
         Height          =   315
         Left            =   -72105
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   12
         Tag             =   "Det"
         Top             =   3060
         Width           =   870
      End
      Begin VB.TextBox txtRFC 
         Height          =   285
         Left            =   -73785
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "Det"
         Top             =   3540
         Width           =   2595
      End
      Begin VB.Frame Frame3 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -71160
         TabIndex        =   9
         Top             =   3180
         Width           =   1695
         Begin FPSpread.vaSpread spdcostos 
            Height          =   2175
            Left            =   120
            OleObjectBlob   =   "frmAtencionClientes.frx":0A57
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtMail 
         Height          =   285
         Left            =   -73785
         TabIndex        =   8
         Tag             =   "Det"
         Top             =   4620
         Width           =   2595
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tesorero:"
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
         Left            =   -69960
         TabIndex        =   43
         Top             =   2340
         Width           =   825
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nóminas:"
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
         Left            =   -69960
         TabIndex        =   42
         Top             =   2700
         Width           =   765
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73800
         TabIndex        =   38
         Top             =   5460
         Width           =   2175
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R Patronal"
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
         Left            =   -74880
         TabIndex        =   37
         Top             =   4020
         Width           =   825
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apoderado:"
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
         Left            =   -69960
         TabIndex        =   35
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Gerente:"
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
         Left            =   -69960
         TabIndex        =   34
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Contador:"
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
         Left            =   -69960
         TabIndex        =   33
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mail Factura:"
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
         Left            =   -74880
         TabIndex        =   29
         Top             =   5055
         Width           =   1020
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   4380
         Width           =   870
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
         Left            =   -74880
         TabIndex        =   26
         Top             =   840
         Width           =   915
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
         Left            =   -74880
         TabIndex        =   25
         Top             =   2175
         Width           =   1020
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
         Left            =   -74880
         TabIndex        =   24
         Top             =   1215
         Width           =   780
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   1740
         Width           =   1005
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
         Left            =   -74880
         TabIndex        =   22
         Top             =   2610
         Width           =   1200
      End
      Begin VB.Label lblPoblacion 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73785
         TabIndex        =   21
         Tag             =   "Det"
         Top             =   2145
         Width           =   3675
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   3090
         Width           =   825
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
         Left            =   -74880
         TabIndex        =   19
         Top             =   3660
         Width           =   360
      End
      Begin VB.Label Label2 
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
         Left            =   -74880
         TabIndex        =   18
         Top             =   4695
         Width           =   540
      End
   End
   Begin VB.Frame fraPrincipal 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Tag             =   "Key"
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Det"
         Top             =   480
         Width           =   7860
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   1200
         Picture         =   "frmAtencionClientes.frx":0CC9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   400
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   1680
         Picture         =   "frmAtencionClientes.frx":0DCB
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   400
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
         TabIndex        =   6
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAtencionClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

