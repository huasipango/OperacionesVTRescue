VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAjustes 
   Caption         =   "Ajustes"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   45
      Top             =   240
      Width           =   6135
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmAjustes.frx":0000
         Left            =   1800
         List            =   "frmAjustes.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   4095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame frmModFecha 
      Caption         =   "Modificar Fecha de Ajuste"
      Height          =   2535
      Left            =   2640
      TabIndex        =   31
      Top             =   4440
      Width           =   5055
      Begin VB.CommandButton cmdCancCambio 
         Caption         =   "Cancelar"
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
         Left            =   2640
         TabIndex        =   39
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdAcepCambio 
         Caption         =   "Cambiar"
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
         Left            =   1200
         TabIndex        =   38
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtFolioCam 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   36
         Top             =   360
         Width           =   810
      End
      Begin MSMask.MaskEdBox mskFechaDispAnt 
         Height          =   345
         Left            =   2520
         TabIndex        =   32
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox mskFechaDispNva 
         Height          =   345
         Left            =   2520
         TabIndex        =   34
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1320
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1860
         TabIndex        =   37
         Top             =   435
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nueva Fecha Disp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   765
         TabIndex        =   35
         Top             =   1395
         Width           =   1545
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Disp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1320
         TabIndex        =   33
         Top             =   915
         Width           =   975
      End
   End
   Begin VB.Frame fraDisp 
      Height          =   4215
      Left            =   5160
      TabIndex        =   23
      Top             =   3120
      Width           =   4935
      Begin FPSpread.vaSpread spdDisp 
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "frmAjustes.frx":001C
         TabIndex        =   24
         Top             =   780
         Width           =   4575
      End
      Begin VB.TextBox txtImpDisp 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   1170
      End
      Begin VB.CommandButton cmdSubeDisp 
         Height          =   330
         Left            =   2160
         Picture         =   "frmAjustes.frx":031C
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "Det"
         ToolTipText     =   "Subir Archivo"
         Top             =   3780
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Importe Dispersion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   375
         Width           =   1695
      End
   End
   Begin VB.Frame fraAju 
      Height          =   4215
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   4935
      Begin FPSpread.vaSpread spdAjustes 
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "frmAjustes.frx":041E
         TabIndex        =   19
         Top             =   780
         Width           =   4575
      End
      Begin VB.CommandButton cmdSubeAju 
         Height          =   330
         Left            =   1800
         Picture         =   "frmAjustes.frx":071E
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "Det"
         ToolTipText     =   "Subir Archivo"
         Top             =   3780
         Width           =   930
      End
      Begin VB.TextBox txtImpAju 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Importe Ajustes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   22
         Top             =   375
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   7320
      Width           =   9975
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   570
         Left            =   6480
         Picture         =   "frmAjustes.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Eliminar"
         Top             =   360
         Width           =   1170
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Concepto"
         Height          =   570
         Left            =   360
         Picture         =   "frmAjustes.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   47
         Tag             =   "Det"
         ToolTipText     =   "Grabar Ajuste"
         Top             =   360
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   570
         Left            =   2400
         Picture         =   "frmAjustes.frx":0A24
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "Det"
         ToolTipText     =   "Grabar Ajuste"
         Top             =   360
         Width           =   1170
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   570
         Left            =   4530
         Picture         =   "frmAjustes.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "Det"
         ToolTipText     =   "Nuevo Ajuste"
         Top             =   360
         Width           =   1170
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   570
         Left            =   8400
         Picture         =   "frmAjustes.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9975
      Begin VB.CommandButton cmdEjecuta 
         Caption         =   "Generar"
         Height          =   570
         Left            =   7920
         Picture         =   "frmAjustes.frx":0D9A
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Salir"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   8520
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frFecha 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   6855
         Begin VB.CommandButton cmdCargaCuentas 
            Caption         =   "Cargar cuentas"
            Height          =   300
            Left            =   4200
            TabIndex        =   43
            Top             =   0
            Width           =   2535
         End
         Begin MSMask.MaskEdBox mskFechaini 
            Height          =   345
            Left            =   1200
            TabIndex        =   8
            Tag             =   "Enc"
            ToolTipText     =   "Fecha del Movimiento"
            Top             =   0
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
         Begin MSMask.MaskEdBox mskFechaFin 
            Height          =   345
            Left            =   2760
            TabIndex        =   9
            Tag             =   "Enc"
            ToolTipText     =   "Fecha del Movimiento"
            Top             =   0
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
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "A:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2400
            TabIndex        =   42
            Top             =   0
            Width           =   165
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "De:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar Fecha de Ajuste"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtFolio 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   810
      End
      Begin VB.ComboBox cboConceptos 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   4410
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   330
         Left            =   6600
         Picture         =   "frmAjustes.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   345
         Left            =   8520
         TabIndex        =   10
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   720
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
      Begin MSMask.MaskEdBox mskFechaDisp 
         Height          =   345
         Left            =   8520
         TabIndex        =   28
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   240
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Disp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7320
         TabIndex        =   29
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   320
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7800
         TabIndex        =   11
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label lblCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   645
      End
   End
   Begin MSComDlg.CommonDialog cmnAbrir 
      Left            =   -120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte
Dim Nombre As String
Public ImpAj10 As Double
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAjustesBE.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(txtFolio.Text)
    mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaDisp.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaDisp.Text, "mm/dd/yyyy")
    producto_cual
    mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub
Sub InicializaForma()
    LimpiarControles Me
    mskFecha.Text = Date
    mskFechaDisp.Text = Date
    txtFolio.Text = BuscaFolio
    cboConceptos.Clear
    txtCliente.Text = ""
    txtNombre.Text = ""
    CargaClaves cboConceptos, "AjustesBE", "Concepto"
    txtCliente.Text = ""
    txtNombre.Text = ""
    spdAjustes.Col = -1
    spdAjustes.Row = -1
    spdAjustes.Action = 12
    spdAjustes.MaxRows = 1
    spddisp.Col = -1
    spddisp.Row = -1
    spddisp.Action = 12
    spddisp.MaxRows = 1
    fraAju.Enabled = False
    fraDisp.Enabled = False
    frmModFecha.Visible = False
    frFecha.Visible = False
    mskFechaIni = Date
    mskFechaFin = Date
    cmdEjecuta.Visible = False
End Sub
Function BuscaFolio()
sqls = " exec Sp_Folio_Sel_Upd 'SEL', 0, 'AJU'"
Set rsBD3 = New ADODB.Recordset
rsBD3.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
If Not rsBD3.EOF Then
    BuscaFolio = rsBD3!Folio
Else
    BuscaFolio = 1
End If
End Function
Private Function SubearchivoDisp() As String
Dim nArchivo, nError, clinea As String, i As Long
Dim valor As Double
Dim Cuenta As Long, empleado As String
Dim noempleado As String
Dim importe As Double
Dim VALIDA As Boolean
Dim ArrVal() As String
Dim cabecera As Boolean
nArchivo = "c:\Facturacion\LogError.txt"
Close #1
Open nArchivo For Output As #1

With spddisp
    Subearchivo = ""
    valor = 0
    On Error GoTo ErrorImport
    cmnAbrir.ShowOpen
    If cmnAbrir.Filename <> "" Then
        nArchivo = FreeFile
        Open cmnAbrir.Filename For Input Access Read As #nArchivo
        i = 0
        sqls = " delete tmppedidobe where usuario = '" & gstrUsuario & "'"
        cnxbdMty.Execute sqls, intRegistros
        MsgBar "Procesando Archivo...", False
        Screen.MousePointer = 11
        Do While Not EOF(nArchivo)
            DoEvents
            Line Input #nArchivo, clinea
            If Not Mid(clinea, 1, 2) = "06" Then
                lblNoEmp = Val(Mid(clinea, 40, 7))
                txtSub = Format(CDbl(Mid(clinea, 47, 12)), "########.00")
                cabecera = True
            Else
                i = i + 1
                'lblNoEmp.Caption = i
                cliente = Val(Mid(clinea, 3, 5))
                noempleado = QuitaCeros(Mid(clinea, 23, 10))
                Cuenta = Mid(clinea, 8, 8)
                Nombre = Trim(Mid(clinea, 33, 26))
                cons = Mid(clinea, 16, 7)
                importe = Val(Mid(clinea, 60, 10))
                cabecera = False
            End If
          
            If cabecera = False Then
                If Cuenta = 0 Then ' si no saben el numero de cuenta lo suben con 0 y se busca
                    'prod = IIf(Product = 8, 6, Product)
                    producto_cual
                    sqls = "sp_CuentasBE_Varios 0," & Product & ",'Busca','" & noempleado & "'," & Val(cliente)
                    
'                    sqls = " select nocuenta, NOMBRE from cuentasbe where empleadora  = " & Val(Cliente) & _
'                           " and noempleado = '" & noempleado & "' and isnull(status,1)  = 1" & _
'                           " and Producto=" & Product & "order by nocuenta desc"
                           
                    Set rsBD = New ADODB.Recordset
                    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rsBD.EOF Then
                        Cuenta = rsBD!noCuenta
                        Nombre = rsBD!Nombre
                    Else
                        If Val(cliente) <> Val(txtCliente.Text) Then
                            Print #1, "El empleado " & noempledao & " no pertenece al cliente " & txtCliente.Text & ""
                        Else
                            Print #1, "Empleado " & noempleado & " no esta dado de alta en el sistema"
                        End If
                        Cuenta = 0
                    End If
                Else  'Si trae numero de cuenta busca si existe
                    'prod = IIf(Product = 8, 6, Product)
                    producto_cual
                    sqls = "sp_CuentasBE_Varios " & Cuenta & "," & Product & ",'Cuenta'"
                    
'                    sqls = " select empleadora,nocuenta, noempleado, nombre, status from cuentasbe where nocuenta = " & cuenta
'                    sqls = sqls & " and Producto=" & Product
                    '  SQLS = SQLS & " and noempleado = '" & noempleado & "'"
                    
                    Set rsBD = New ADODB.Recordset
                    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                    If Not rsBD.EOF Then
                        If rsBD!Status = 2 Then
                            Print #1, "Empleado " & noempleado & " : La cuenta esta cancelada en el sistema"
                            Cuenta = 0
                        ElseIf Val(rsBD!Empleadora) <> Val(txtCliente.Text) And cboConceptos.ListIndex <> 2 Then
                            Print #1, "Empleado " & noempleado & " no pertenece a la empleadora " & rsBD!Empleadora & ""
                            Cuenta = 0
                        ElseIf UCase(Trim(rsBD!noempleado)) <> UCase(Trim(noempleado)) And UCase(Trim(rsBD!NoEmpleadoNvo)) <> UCase(Trim(noempleado)) Then
                            Print #1, "Empleado " & noempleado & " : El empleado de la cuenta es diferente al empleado del archivo"
                            Cuenta = 0
                        Else
                            Cuenta = rsBD!noCuenta
                            Nombre = rsBD!Nombre
                        End If
                    Else
                        Print #1, "Empleado " & noempleado & " : La cuenta no esta dada de alta en el sistema"
                        Cuenta = 0
                    End If
    
                End If
                
                
                If importe = 0 Then
                     Print #1, "Empleado " & noempleado & " : El importe a depositar es 0"
                End If
                    
                If Cuenta <> 0 And importe <> 0 Then
                    sqls = " exec sp_tmpPedidosBE @usuario = '" & gstrUsuario & "'" & _
                           " , @sucursal = " & cliente & _
                           " , @Cuenta =  " & Cuenta & _
                           " , @Empleado = '" & noempleado & "'" & _
                           " , @Nombre   = '" & Nombre & "'" & _
                           " , @Valor   = " & importe & _
                           " , @Producto=" & Product
                           
                    cnxbdMty.Execute sqls, intRegistros
                    
                End If
            End If
        Loop
        
        .Col = -1
        .Row = -1
        .Action = 12
        MsgBar "Leyendo Registros...", False
        
        sqls = "sp_CuentasBE_Varios 0," & Product & ",'Temporal','" & gstrUsuario & "'"
        
'        sqls = " select  usuario, sucursal, cuenta, empleado, Nombre, sum(valor) Total" & _
'               " From tmppedidobe " & _
'               " where usuario = '" & gstrUsuario & "' and Producto=" & Product & _
'               " group by usuario, sucursal, cuenta, empleado, Nombre" & _
'               " order by  cuenta"
        
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        i = 0
        valor = 0
        
        Do While Not rsBD.EOF
            Screen.MousePointer = 11
            i = i + 1
            .MaxRows = i
            .Row = i
            .Col = 1
            .Text = rsBD!Cuenta
            .Col = 2
            .Text = rsBD!empleado
            .Col = 3
            .Text = rsBD!Nombre
            
            .Col = 4
            .Text = CDbl(rsBD!total)
            
            valor = valor + CDbl(rsBD!total)
            rsBD.MoveNext
        Loop
        
        txtImpDisp.Text = Format(CDbl(valor), "########.00")
        
        Close #nArchivo
        Close #1
        Screen.MousePointer = 1
   
        MsgBar "Listo...", False
        
        Screen.MousePointer = 1
        
        If FileLen("C:\facturacion\LogError.TXT") > 0 Then
            RetVal = Shell("C:\WINDOWS\NOTEPAD.EXE C:\facturacion\LogError.TXT", 1)
        End If
        
        MsgBox "El archivo " & cmnAbrir.Filename & " se importó exitosamente, con " _
            & i & " registros", vbInformation + vbOKOnly, Me.Caption
            
            
        Subearchivo = True
        cmnAbrir.Filename = ""
        Exit Function
    Else
        Exit Function
    End If
End With

rsBD.Close
Set rsBD = Nothing
Close #1
ErrorImport:
    Beep
    MsgBox "Hubo un error al actualizar! Favor de avisar a sistemas! Error: " & ERR.Number & vbCrLf & ERR.Description, vbCritical + vbOKOnly, Me.Caption
    Screen.MousePointer = 1
    Resume Next
End Function
Private Function SubearchivoAju() As String
Dim nArchivo, nError, clinea As String, i As Long
Dim valor As Double
Dim Cuenta As Long, empleado As String
Dim noempleado As String
Dim importe As Double
Dim VALIDA As Boolean
Dim ArrVal() As String
Dim cabecera As Boolean
nArchivo = "c:\Facturacion\LogError.txt"
Close #1
Open nArchivo For Output As #1

With spdAjustes
    Subearchivo = ""
    valor = 0
    On Error GoTo ErrorImport
    cmnAbrir.ShowOpen
    If cmnAbrir.Filename <> "" Then
        nArchivo = FreeFile
        Open cmnAbrir.Filename For Input Access Read As #nArchivo
        i = 0
        sqls = "sp_CuentasBE_Varios 0," & Product & ",'BorraTemp','" & gstrUsuario & "'"
'        sqls = " delete tmppedidobe where usuario = '" & gstrUsuario & "'"
'        sqls = sqls & " and Producto=" & Product
        
        cnxbdMty.Execute sqls, intRegistros
        MsgBar "Procesando Archivo...", False
        Screen.MousePointer = 11
        Do While Not EOF(nArchivo)
            DoEvents
            Line Input #nArchivo, clinea
            If Not Mid(clinea, 1, 2) = "06" Then
                lblNoEmp = Val(Mid(clinea, 40, 7))
                txtSub = Format(CDbl(Mid(clinea, 47, 12)), "########.00")
                cabecera = True
            Else
                i = i + 1
                'lblNoEmp.Caption = i
                cliente = Val(Mid(clinea, 3, 5))
                noempleado = QuitaCeros(Mid(clinea, 23, 10))
                Cuenta = Mid(clinea, 8, 8)
                Nombre = Trim(Mid(clinea, 33, 26))
                cons = Mid(clinea, 16, 7)
                importe = Val(Mid(clinea, 60, 10))
                cabecera = False
            End If
          
            If cabecera = False Then
                If Cuenta = 0 Then ' si no saben el numero de cuenta lo suben con 0 y se busca
                    'prod = IIf(Product = 8, 6, Product)
                    producto_cual
                    sqls = "sp_CuentasBE_Varios 0," & Product & ",'Busca','" & noempleado & "'," & Val(cliente)
                    
'                    sqls = " select nocuenta, NOMBRE from cuentasbe where empleadora  = " & Val(Cliente) & _
'                           " and noempleado = '" & noempleado & "' and isnull(status,1)  = 1" & _
'                           " and Producto=" & Product & " order by nocuenta desc"
                           
                    Set rsBD = New ADODB.Recordset
                    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rsBD.EOF Then
                        Cuenta = rsBD!noCuenta
                        Nombre = rsBD!Nombre
                    Else
                        If Val(cliente) <> Val(txtCliente.Text) Then
                            Print #1, "El empleado " & noempledao & " no pertenece al cliente " & txtCliente.Text & ""
                        Else
                            Print #1, "Empleado " & noempleado & " no esta dado de alta en el sistema"
                        End If
                        Cuenta = 0
                    End If
                Else  'Si trae numero de cuenta busca si existe
                    'prod = IIf(Product = 8, 6, Product)
                    producto_cual
                    sqls = "sp_CuentasBE_Varios " & Cuenta & "," & Product & ",'Cuenta'"
                    
'                    sqls = " select empleadora,nocuenta, noempleado, nombre, status from cuentasbe where nocuenta = " & cuenta
'                    sqls = sqls & " and Producto=" & Product
                    Set rsBD = New ADODB.Recordset
                    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                    If Not rsBD.EOF Then
                        If rsBD!Status = 2 And cboConceptos.ItemData(cboConceptos.ListIndex) <> 10 Then
                            Print #1, "Empleado " & noempleado & " : La cuenta esta cancelada en el sistema"
                            Cuenta = 0
                        ElseIf Val(rsBD!Empleadora) <> Val(txtCliente.Text) Then
                            Print #1, "Empleado " & noempleado & " no pertenece a la empleadora " & rsBD!Empleadora & ""
                            Cuenta = 0
                        ElseIf UCase(Trim(rsBD!noempleado)) <> UCase(Trim(noempleado)) And UCase(Trim(rsBD!NoEmpleadoNvo)) <> UCase(Trim(noempleado)) Then
                            Print #1, "Empleado " & noempleado & " : El empleado de la cuenta es diferente al empleado del archivo"
                            Cuenta = 0
                        Else
                            Cuenta = rsBD!noCuenta
                            Nombre = rsBD!Nombre
                        End If
                    Else
                        Print #1, "Empleado " & noempleado & " : La cuenta no esta dada de alta en el sistema"
                        Cuenta = 0
                    End If
    
                End If
                
                
                If importe = 0 Then
                     Print #1, "Empleado " & noempleado & " : El importe a depositar es 0"
                End If
                    
                If Cuenta <> 0 And importe <> 0 Then
                    sqls = " exec sp_tmpPedidosBE @usuario = '" & gstrUsuario & "'" & _
                           " , @sucursal = " & cliente & _
                           " , @Cuenta =  " & Cuenta & _
                           " , @Empleado = '" & noempleado & "'" & _
                           " , @Nombre   = '" & Nombre & "'" & _
                           " , @Valor   = " & importe & _
                           " , @Producto=" & Product
                           
                    cnxbdMty.Execute sqls, intRegistros
                    
                End If
            End If
        Loop
        
        .Col = -1
        .Row = -1
        .Action = 12
        MsgBar "Leyendo Registros...", False
        
        sqls = "sp_CuentasBE_Varios 0," & Product & ",'Temporal','" & gstrUsuario & "'"
        
'        sqls = " select  usuario, sucursal, cuenta, empleado, Nombre, sum(valor) Total" & _
'               " From tmppedidobe " & _
'               " where usuario = '" & gstrUsuario & "'" & _
'               " and Producto=" & Product & _
'               " group by usuario, sucursal, cuenta, empleado, Nombre" & _
'               " order by  cuenta"
        
        
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        i = 0
        valor = 0
        
        Do While Not rsBD.EOF
            Screen.MousePointer = 11
            i = i + 1
            .MaxRows = i
            .Row = i
            .Col = 1
            .Text = rsBD!Cuenta
            .Col = 2
            .Text = rsBD!empleado
            .Col = 3
            .Text = rsBD!Nombre
            
            .Col = 4
            .Text = CDbl(rsBD!total)
            
            valor = valor + CDbl(rsBD!total)
            rsBD.MoveNext
        Loop
        
        txtImpAju.Text = Format(CDbl(valor), "########.00")
        
        Close #nArchivo
        Close #1
        Screen.MousePointer = 1
   
        MsgBar "Listo...", False
        
        Screen.MousePointer = 1
        
        If FileLen("C:\facturacion\LogError.TXT") > 0 Then
            RetVal = Shell("C:\WINDOWS\NOTEPAD.EXE C:\facturacion\LogError.TXT", 1)
        End If
        
        MsgBox "El archivo " & cmnAbrir.Filename & " se importó exitosamente, con " _
            & i & " registros", vbInformation + vbOKOnly, Me.Caption
            
        Subearchivo = True
        cmnAbrir.Filename = ""
        Exit Function
    Else
        Exit Function
    End If
End With

rsBD.Close
Set rsBD = Nothing
Close #1
ErrorImport:
    Beep
    MsgBox "Hubo un error al actualizar! Favor de avisar a sistemas! Error: " & ERR.Number & vbCrLf & ERR.Description, vbCritical + vbOKOnly, Me.Caption
    Screen.MousePointer = 1
    Resume Next
End Function

Private Sub cboConceptos_Click()
Dim interno As Boolean
frFecha.Visible = False

If cboConceptos.ItemData(cboConceptos.ListIndex) = 3 And txtCliente.Text = "" Then
   MsgBox "En una nota de consumo primero debe seleccionar el cliente interno", vbExclamation, "No ha seleccionado cliente interno aun"
   txtCliente.SetFocus
   InicializaForma
   Exit Sub
End If

sql = "SELECT * from Clientes where Cliente=" & Val(txtCliente.Text)
sql = sql & " And RFC=''"
Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
   interno = True
Else
   interno = False
End If

Select Case cboConceptos.ItemData(cboConceptos.ListIndex)
    Case 1, 10
       fraAju.Enabled = True
       fraDisp.Enabled = True
       fraAju.Visible = True
       fraDisp.Visible = True
       cmdEjecuta.Visible = False
       
    Case 2, 4, 5, 7, 8, 9, 11, 12
       fraAju.Enabled = True
       fraDisp.Enabled = False
       fraAju.Visible = True
       fraDisp.Visible = False
       cmdEjecuta.Visible = False
       If cboConceptos.ItemData(cboConceptos.ListIndex) = 12 Then
          frFecha.Visible = True
          cmdEjecuta.Visible = True
       End If
    Case 3, 6, 13
       If cboConceptos.ItemData(cboConceptos.ListIndex) = 3 And interno = False Then
          MsgBox "Lo siento pero el cliente proporcionado no es cliente interno", vbCritical, "Cliente no es interno para notas de consumo"
          InicializaForma
          Exit Sub
       End If
       fraAju.Enabled = False
       fraDisp.Enabled = True
       fraAju.Visible = False
       fraDisp.Visible = True
       cmdEjecuta.Visible = False
End Select
End Sub

Private Sub cmdAcepCambio_Click()

If mskFechaDispNva.Text = "__/__/____" Then
    MsgBox "Primero debe capturar la fecha nueva de dispersion", vbCritical
    Exit Sub
End If

If IsDate(mskFechaDispNva.Text) Then
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = " update ajustesbe set fechamov = '" & Format(mskFechaDispNva.Text, "mm/dd/yyyy") & "'" & _
           " where folio = " & Val(txtFolioCam) & " AND Producto=" & Product
         
    cnxbdMty.Execute sqls, intRegistros
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = " update ajustesdetbe set fechadisp = '" & Format(mskFechaDispNva.Text, "mm/dd/yyyy") & "'" & _
           " where folio = " & Val(txtFolioCam) & " and Producto=" & Product
         
    cnxbdMty.Execute sqls, intRegistros
    
    MsgBox "Fecha Actualizada !!", vbInformation, "Fecha actualizada"
    
    txtFolioCam.Text = ""
    mskFechaDispNva.Text = "__/__/____"
    mskFechaDispAnt.Text = "__/__/____"
    txtFolioCam.SetFocus

Else
    MsgBox "Fecha inválida!!", vbCritical, "Fecha Invalida"
    mskFechaDispNva.Text = "__/__/____"
    mskFechaDispNva.SetFocus
    Exit Sub
End If
End Sub

Private Sub CmdBuscar_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "Cliente"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtCliente = frmConsulta.cliente
       txtNombre = frmConsulta.Nombre
      
    End If
    Set frmConsulta = Nothing
    MsgBar "", False

End Sub

Private Sub cmdCancCambio_Click()
   frmModFecha.Visible = False
End Sub

Private Sub cmdCancelar_Click()
  ajuste_o_cancel = 1
  frmConceptos.Show 1
End Sub

' PENDIENTE CREAR TABLA LIQUIDACIONESPG --ya no

Private Sub cmdCargaCuentas_Click()
If txtFolio <> "" And txtCliente <> "" Then
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = " select a.cuenta, b.noempleado empleado, b.nombre, count(a.importe) * 1.74 Total" & _
           " from liquidacionesbe a (nolock), cuentasbe b (nolock)" & _
           " where a.fechatran between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'" & _
           " and  a.cuenta = b.nocuenta" & _
           " and  a.cveresp = 51" & _
           " and a.cliente = " & txtCliente.Text & _
           " and isnull(b.status,1) <> 2 " & _
           " and b.Producto=" & Product & _
           " group by a.cuenta, b.noempleado,b.nombre" & _
           " order by count(a.importe) desc"
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    With spdAjustes
    .Col = -1
    .Row = -1
    .Action = 12
    i = 0
      Do While Not rsBD.EOF
            Screen.MousePointer = 11
            i = i + 1
            .MaxRows = i
            .Row = i
            .Col = 1
            .Text = rsBD!Cuenta
            .Col = 2
            .Text = rsBD!empleado
            .Col = 3
            .Text = rsBD!Nombre
            
            .Col = 4
            .Text = CDbl(rsBD!total)
            
            valor = valor + CDbl(rsBD!total)
            rsBD.MoveNext
           
        Loop
    End With
    
    txtImpAju.Text = Format(CDbl(valor), "########.00")
    Screen.MousePointer = 1
Else
  MsgBox "Faltan datos por capturar...", vbExclamation, "Datos Faltantes"
  Exit Sub
End If
End Sub

Private Sub cmdEjecuta_Click()
Dim folic As Long, i As Integer, cliente As Long, tot As Double
On Error GoTo ERR
Me.Caption = "Ajustes"
If MsgBox("La siguiente accion se tardará aproximadamente 5 minutos(+/-);una vez iniciado no se detendra y se crearán consecutivos de ajustes automaticamente" & vbCrLf & "¿Esta seguro de realizar esta accion ahora?", vbQuestion + vbYesNo + vbDefaultButton2, "Generador de ajustes x fondos insufucientes") = vbYes Then
   rsBD.Close
   Set rsBD = Nothing
   Me.Caption = "Calculando ajustes de Clientes"
    sqls = "exec spr_AjustesFondosInsuf_sl  @Producto = " & Product & _
                                     " , @FechaIni = '" & Format(mskFechaIni, "mm/dd/yyyy") & "'" & _
                                     " , @FechaFin = '" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'"
                                
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   Do While Not rsBD.EOF
      cliente = Val(rsBD!cliente)
      tot = Val(rsBD!total)
      folic = BuscaFolio
      Me.Caption = "Procesando ajuste del Cliente " & Val(cliente)
      sqls = " exec sp_AjustesBe @Folio = " & Val(folic) & _
                           " , @Cliente = " & Val(cliente) & _
                           " , @concepto=  " & cboConceptos.ItemData(cboConceptos.ListIndex) & _
                           " , @Cargo = " & Val(tot) & _
                           " , @Abono = " & Val(0) & _
                           " , @Usuario   = '" & gstrUsuario & "'" & _
                           " , @Fecha   =  '" & Format(Date + 1, "MM/DD/YYYY") & "'" & _
                           " , @Producto=" & Product
      cnxbdMty.Execute sqls, intRegistros
      
      '---detalle
      sqls = "exec spr_AjustesFondosInsufDet_sl   @Producto = " & Product & _
                                            " , @FechaIni = '" & Format(mskFechaIni, "mm/dd/yyyy") & "'" & _
                                            " , @FechaFin = '" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'" & _
                                            " , @Cliente  = " & Val(cliente)

      Set rsBD2 = New ADODB.Recordset
      rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
       
      '----graba detalle
       If Not rsBD2.EOF Then
          i = 1
          Do While Not rsBD2.EOF
             sqls = " exec sp_AjustesDetBe @Folio = " & Val(folic) & _
                   " , @Refer = " & i & _
                   " , @Cliente = " & cliente & _
                   " , @Tipomov=  'C'" & _
                   " , @Cuenta = " & rsBD2!Cuenta & _
                   " , @Importe = " & rsBD2!total & _
                   " , @Concepto  = " & cboConceptos.ItemData(cboConceptos.ListIndex) & _
                   " , @FechaDisp      =    '" & Format(Date + 1, "MM/DD/YYYY") & "'" & _
                   " , @Producto=" & Product
                   cnxbdMty.Execute sqls, intRegistros
                   i = i + 1
                   rsBD2.MoveNext
          Loop
          sqls = " exec Sp_Folio_Sel_Upd 'UPD',0, 'AJU',  " & Val(folic) & ""
          cnxbdMty.Execute sqls, intRegistros
       End If
       '----------
       rsBD.MoveNext
   Loop
   MsgBox "Se terminaron de generar todos los ajustes x fondos insuficientes automaticamente", vbInformation, "Hecho..."
   Me.Caption = "Ajustes"
End If
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub

Private Sub cmdGrabar_Click()
frmSecreto.Show 1
If palabra_ok = True Then
With spdAjustes
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        Cuenta = Val(.Text)
        If Cuenta = 0 Then
            .Row = i
            .Action = 5
            .MaxRows = .MaxRows - 1
        End If
    Next i
End With

With spddisp
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        Cuenta = Val(.Text)
        If Cuenta = 0 Then
            .Row = i
            .Action = 5
            .MaxRows = .MaxRows - 1
        End If
    Next i
End With

If ValidaDatos Then
     If GrabaAjuste Then
        GrabaAjusteDet
        sqls = " exec Sp_Folio_Sel_Upd 'UPD',0, 'AJU',  " & Val(txtFolio.Text) & ""
        cnxbdMty.Execute sqls, intRegistros
     End If
    
     MsgBox "Datos Actualizados!", vbInformation, "Ok..."
     Imprime (crptToWindow)
     InicializaForma
End If
End If
End Sub
Sub GeneraArchivo(tipo As String)
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre
Dim Empleadora As String
Dim RetVal

On Error GoTo err_gral
               
If tipo = "Ajuste" Then

    Open "c:\Facturacion\Ajuste" & txtCliente & Format(Date, "DDMM") & ".txt" For Output As #1

    clinea = "051003602Vale Total S.A. DE C.V.       "
    If Product = 1 Then
       clinea = clinea & Format(spdAjustes.MaxRows, "0000000") & Format(txtImpAju.Text, "000000000.00") & "Despensa" & Format(txtFecha, "DD-MMM-YY")
    ElseIf Product = 2 Then
       clinea = clinea & Format(spdAjustes.MaxRows, "0000000") & Format(txtImpAju.Text, "000000000.00") & "Combustible" & Format(txtFecha, "DD-MMM-YY")
    ElseIf Product = 4 Then
       clinea = clinea & Format(spdAjustes.MaxRows, "0000000") & Format(txtImpAju.Text, "000000000.00") & "Viaticos" & Format(txtFecha, "DD-MMM-YY")
    End If
    Print #1, clinea

    With spdAjustes
         For i = 1 To .MaxRows
               clinea = "06"
               cliente = Val(txtCliente.Text)
               clinea = clinea & Format(Val(.Text), "00000") '& Space(5) & Format(i, "0000000")
               .Row = i
               .Col = 1
               Cuenta = .Text
               clinea = clinea & Format(Val(Cuenta), "00000000")
               clinea = clinea & Format(i, "0000000")
               .Col = 2
               clinea = clinea & Pad(.Text, 10, "0", "L")
               .Col = 3
               clinea = clinea & Pad(Trim(.Text), 26, " ", "R")
               .Col = 4
               clinea = clinea & Format(CDbl(.Text), "0000000.00")
               Print #1, clinea
         Next i
         Close #1
         
         If FileLen("C:\Facturacion\Ajuste" & txtCliente & Format(Date, "DDMM") & ".txt") > 0 Then
            RetVal = Shell("C:\WINDOWS\NOTEPAD.EXE c:\Facturacion\Ajuste" & txtCliente & Format(Date, "DDMM") & ".txt", 1)
         End If
        
    End With
ElseIf tipo = "Dispersion" Then
    Open "c:\Facturacion\Disp" & txtCliente & Format(Date, "DDMM") & ".txt" For Output As #1

    clinea = "051003602Vale Total S.A. DE C.V.       "
    If Product = 1 Then
       clinea = clinea & Format(spddisp.MaxRows, "0000000") & Format(txtImpDisp.Text, "000000000.00") & "Despensa" & Format(txtFecha, "DD-MMM-YY")
    ElseIf Product = 2 Then
       clinea = clinea & Format(spddisp.MaxRows, "0000000") & Format(txtImpDisp.Text, "000000000.00") & "Combustible" & Format(txtFecha, "DD-MMM-YY")
    ElseIf Product = 4 Then
       clinea = clinea & Format(spddisp.MaxRows, "0000000") & Format(txtImpDisp.Text, "000000000.00") & "Viaticos" & Format(txtFecha, "DD-MMM-YY")
    End If
    Print #1, clinea

    With spddisp
         For i = 1 To .MaxRows
               clinea = "06"
               cliente = Val(txtCliente.Text)
               clinea = clinea & Format(Val(.Text), "00000") '& Space(5) & Format(i, "0000000")
               .Row = i
               .Col = 1
               Cuenta = .Text
               clinea = clinea & Format(Val(Cuenta), "00000000")
               clinea = clinea & Format(i, "0000000")
               .Col = 2
               clinea = clinea & Pad(.Text, 10, "0", "L")
               .Col = 3
               clinea = clinea & Pad(Trim(.Text), 26, " ", "R")
               .Col = 4
               clinea = clinea & Format(CDbl(.Text), "0000000.00")
               Print #1, clinea
         Next i
         Close #1
         
         If FileLen("C:\Facturacion\Disp" & txtCliente & Format(Date, "DDMM") & ".txt") > 0 Then
            RetVal = Shell("C:\WINDOWS\NOTEPAD.EXE c:\Facturacion\Disp" & txtCliente & Format(Date, "DDMM") & ".txt", 1)
         End If
    End With
End If
Exit Sub
err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmAjustes.GeneraArchivo")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, vbExclamation, "Errores generados"
   MsgBar "", False
End Sub
Function GrabaAjuste()
GrabaAjuste = False

On Error GoTo err_graba
  'prod = IIf(Product = 8, 6, Product)
  producto_cual
  sqls = " exec sp_AjustesBe @Folio = " & Val(txtFolio.Text) & _
                           " , @Cliente = " & Val(txtCliente.Text) & _
                           " , @concepto=  " & cboConceptos.ItemData(cboConceptos.ListIndex) & _
                           " , @Cargo = " & Val(txtImpAju.Text) & _
                           " , @Abono = " & Val(txtImpDisp.Text) & _
                           " , @Usuario   = '" & gstrUsuario & "'" & _
                           " , @Fecha   =  '" & Format(mskFechaDisp.Text, "MM/DD/YYYY") & "'" & _
                           " , @Producto=" & Product
  cnxbdMty.Execute sqls, intRegistros
  
  If cboConceptos.ItemData(cboConceptos.ListIndex) = 3 Then 'Notas de Consumo
  
    sqls = "select bodega from clientes where cliente = " & Val(txtCliente.Text)

    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
      Bodega = rsBD!Bodega
    Else
      Bodega = 1
    End If
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "exec spb_GrabaNotaConsumoBE @Sucursal =" & gnBodega & _
           ",@Producto=" & Product & _
           ",@TipoPedido = 11" & _
           ",@Pasan = 0" & _
           ",@Motivo = ''" & _
           ",@TipoFactura = 3" & _
           ",@Cliente = " & Val(txtCliente.Text) & _
           ",@nValorPedido = " & Val(txtImpDisp.Text) & _
           ",@Usuario   = '" & gstrUsuario & "'"
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        Pedido = rsBD!Pedido
        Serie = Trim(rsBD!Serie)
        Factura = rsBD!Factura
    End If
             
End If
GrabaAjuste = True
                                   
Exit Function

err_graba:
        MsgBox ERR.Description, vbCritical, "Errores generados al grabar"
        GrabaAjuste = False
        Resume Next
End Function
Function GrabaAjusteDet()
GrabaAjusteDet = False

On Error GoTo err_graba
  'Graba Ajustes
  Folio = Val(txtFolio.Text)
  cliente = Val(txtCliente.Text)
  
  If cboConceptos.ItemData(cboConceptos.ListIndex) <= 2 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 4 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 5 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 7 Or cboConceptos.ItemData(cboConceptos.ListIndex) >= 8 Then
    With spdAjustes
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            Cuenta = Val(.Text)
            .Col = 4
            importe = Val(.Text)
            
            
            If Cuenta <> 0 And importe <> 0 Then
                'prod = IIf(Product = 8, 6, Product)
                producto_cual
                sqls = " exec sp_AjustesDetBe @Folio = " & Folio & _
                           " , @Refer = " & i & _
                           " , @Cliente = " & cliente & _
                           " , @Tipomov=  'C'" & _
                           " , @Cuenta = " & Cuenta & _
                           " , @Importe = " & importe & _
                           " , @Concepto  = " & cboConceptos.ItemData(cboConceptos.ListIndex) & _
                           " , @FechaDisp      =    '" & Format(mskFechaDisp.Text, "MM/DD/YYYY") & "'" & _
                           " , @Producto=" & Product

                cnxbdMty.Execute sqls, intRegistros
            End If
        Next i
     End With
  End If
    
   
  If cboConceptos.ItemData(cboConceptos.ListIndex) = 1 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 10 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 3 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 6 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 13 Then
  With spddisp
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            Cuenta = Val(.Text)
            .Col = 2
            NumEmp = Trim(.Text)
            .Col = 3
            NomEmp = Trim(.Text)
            .Col = 4
            importe = Val(.Text)
            
            If Cuenta <> 0 And importe <> 0 Then
                'prod = IIf(Product = 8, 6, Product)
                producto_cual
                sqls = " exec sp_AjustesDetBe @Folio = " & Folio & _
                           " , @Refer = " & i & _
                           " , @Cliente = " & cliente & _
                           " , @Tipomov=  'A'" & _
                           " , @Cuenta = " & Cuenta & _
                           " , @Importe = " & importe & _
                           " , @Concepto  = " & cboConceptos.ItemData(cboConceptos.ListIndex) & _
                           " , @FechaDisp      =    '" & Format(mskFechaDisp.Text, "MM/DD/YYYY") & "'" & _
                           " , @Producto=" & Product

                cnxbdMty.Execute sqls, intRegistros
                         
                sqls = " EXEC sp_Recibos_Ins "
                sqls = sqls & vbCr & "  @Sucursal       = 1"
                sqls = sqls & vbCr & ", @Pedido      =" & Folio
                sqls = sqls & vbCr & ", @cliente  =    " & cliente
                sqls = sqls & vbCr & ", @Folio    =    " & i
                sqls = sqls & vbCr & ", @NumEmpl      =    '" & NumEmp & "'"
                sqls = sqls & vbCr & ", @NombreEmpl    =    '" & NomEmp & "'"
                sqls = sqls & vbCr & ", @Numdepto    =    " & cliente
                sqls = sqls & vbCr & ", @Valor    =  " & importe
                sqls = sqls & vbCr & ", @Pagina       = 1"
                sqls = sqls & vbCr & ", @Columna = 0"
                sqls = sqls & vbCr & ", @Renglon   = 0"
                sqls = sqls & vbCr & ", @Campo01    =    '" & Format(mskFechaDisp.Text, "MM/DD/YYYY") & "'"
                sqls = sqls & vbCr & ", @Campo02    = " & Cuenta
                sqls = sqls & vbCr & ", @Campo03    ='A'"
                sqls = sqls & vbCr & ", @Valor1    =   0"
                sqls = sqls & vbCr & ", @Valor3    =   " & cboConceptos.ItemData(cboConceptos.ListIndex)
                
                cnxbdMty.Execute sqls, intRegistros
        
            End If
        Next i
     End With
  End If
  GrabaAjusteDet = True
Exit Function

err_graba:
        MsgBox "Error al grabar el Detalle del ajuste " & ERR.Description, vbCritical, "Errores generados"
        GrabaAjusteDet = False
End Function

Function ValidaDatos()
Dim interno As Boolean
ValidaDatos = True
    
    sql = "SELECT * from Clientes where Cliente=" & Val(txtCliente.Text)
    sql = sql & " And RFC=''"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       interno = True
    Else
       interno = False
    End If
     
    If cboConceptos.ItemData(cboConceptos.ListIndex) = 3 And interno = False Then
          MsgBox "Lo siento pero el cliente proporcionado no es cliente interno", vbCritical, "Cliente no es interno para notas de consumo"
          ValidaDatos = False
          Exit Function
    End If
      
    If Trim(txtCliente.Text) = "" Then
        MsgBox "Falta Capturar Numero de Cliente", vbInformation, "Numero de Cliente pte."
        ValidaDatos = False
        Exit Function
    End If
    
    If cboConceptos.ListIndex < 0 Then
        MsgBox "Falta seleccionar el concepto del ajuste", vbInformation, "Que Concepto?..."
        ValidaDatos = False
        Exit Function
    End If
    
    If cboConceptos.ItemData(cboConceptos.ListIndex) = 1 And (Val(txtImpAju.Text) <> Val(txtImpDisp.Text)) Then
        MsgBox "El importe de los ajustes es diferente al importe de las dispersiones, por favor revise el detalle.", vbInformation, "Importe diferente"
        ValidaDatos = False
        Exit Function
    End If
    
    If cboConceptos.ItemData(cboConceptos.ListIndex) = 1 And (Val(txtImpAju.Text) = 0 And Val(txtImpDisp.Text) = 0) Then
        MsgBox "El importe de los saldos a dispersar o a ajustar no puede ser cero", vbInformation, "No pueden haber valores en ceros"
        ValidaDatos = False
        Exit Function
    End If
    
    If (cboConceptos.ItemData(cboConceptos.ListIndex) = 2 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 4) And Val(txtImpAju.Text) = 0 Then
        MsgBox "El importe de los saldos a ajustar no puede ser cero", vbInformation, "No puede haber importes en ceros"
        ValidaDatos = False
        Exit Function
    End If
    
    If cboConceptos.ItemData(cboConceptos.ListIndex) = 3 And Val(txtImpDisp.Text) = 0 Then
        MsgBox "El importe de los saldos a dispersar no puede ser cero", vbInformation, "No puede haber importes en ceros"
        ValidaDatos = False
        Exit Function
    End If
    
End Function
Private Sub cmdNuevo_Click()
    InicializaForma
End Sub



Private Sub cmdPresentar_Click()

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSubeAju_Click()
    Call SubearchivoAju
End Sub

Sub Descarga_Rechazos()
    If Val(txtCliente.Text) = 0 Then
       MsgBox "El numero de cliente no es valido...verifiquelo", vbCritical, "Error en numero de cliente"
       Exit Sub
    End If
    sqls = "Select s.cuenta,c.noempleado empleado,c.Nombre,s.importe From Sbidispersiones s"
    sqls = sqls & " Inner join liquidacionesbe l (nolock) on l.cuenta=s.cuenta and l.cliente=s.cliente and l.importe<>s.importe"
    sqls = sqls & " Inner join cuentasbe c (nolock) on c.nocuenta=s.cuenta"
    sqls = sqls & " and s.producto=" & Product
    sqls = sqls & " and l.producto=" & Product
    sqls = sqls & " and c.producto=" & Product
    sqls = sqls & " and l.fechatran='" & Format(mskFechaDisp.Text, "mm/dd/yyyy") & "'"
    sqls = sqls & " and s.FechaProc='" & Format(mskFechaDisp.Text, "mm/dd/yyyy") & "'"
    sqls = sqls & " and s.cliente=" & Val(txtCliente.Text)
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    With spddisp
    .Col = -1
    .Row = -1
    .Action = 12
     i = 0
      Do While Not rsBD.EOF
             i = i + 1
            .MaxRows = i
            .Row = i
            .Col = 1
            .Text = rsBD!Cuenta
            .Col = 2
            .Text = rsBD!empleado
            .Col = 3
            .Text = rsBD!Nombre
            
            .Col = 4
            .Text = CDbl(rsBD!importe)
            
            valor = valor + CDbl(rsBD!importe)
            rsBD.MoveNext
        Loop
    End With
    
    txtImpDisp.Text = Format(CDbl(valor), "########.00")
    
End Sub

Private Sub cmdSubeDisp_Click()
   If cboConceptos.ItemData(cboConceptos.ListIndex) = 13 Then
      Call Descarga_Rechazos
   Else
      Call SubearchivoDisp
   End If
End Sub

Private Sub Command1_Click()
frmModFecha.Visible = True

txtFolioCam.Text = ""
mskFechaDispAnt.Text = "__/__/____"
mskFechaDispNva.Text = "__/__/____"
txtFolioCam.SetFocus
End Sub

Private Sub Command2_Click()
Dim j As Integer
    sqls = " select cliente,   sum(importe)Importe" & _
           " From dbo.Hoja1$" & _
           " group by cliente order by cliente"

    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    Folio = 1164
    
    i = 0
    
    Do While Not rsBD.EOF
    Folio = Folio + 1
    j = 0
          'prod = IIf(Product = 8, 6, Product)
          producto_cual
          sqls = " exec sp_AjustesBe @Folio = " & Val(Folio) & _
                           " , @Cliente = " & rsBD!cliente & _
                           " , @concepto=  5" & _
                           " , @Cargo = " & rsBD!importe & _
                           " , @Abono = 0 " & _
                           " , @Usuario   = '71485'" & _
                           " , @Fecha   =  '07/22/2009'" & _
                           " , @Producto=" & Product
                           
        cnxbdMty.Execute sqls, intRegistros
    
        sqls = " select cliente, CUENTA,   sum(importe)Importe" & _
               " From dbo.Hoja1$" & _
               " WHERE CLIENTE = " & rsBD!cliente & _
               " group by cliente , CUENTA order by cliente,cuenta"
    
        Set rsBD2 = New ADODB.Recordset
        rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
            
        Do While Not rsBD2.EOF
            j = j + 1
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = " exec sp_AjustesDetBe @Folio = " & Folio & _
                           " , @Refer = " & j & _
                           " , @Cliente = " & rsBD2!cliente & _
                           " , @Tipomov=  'C'" & _
                           " , @Cuenta = " & rsBD2!Cuenta & _
                           " , @Importe = " & rsBD2!importe & _
                           " , @Concepto  = 5" & _
                           " , @FechaDisp      =  '07/22/2009'" & _
                           " , @Producto=" & Product
                           

            cnxbdMty.Execute sqls, intRegistros
        
        rsBD2.MoveNext
        
        Loop
        
     rsBD.MoveNext
    Loop

End Sub

Private Sub cboProducto_Click()
  Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
  'Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
End Sub

Private Sub Command3_Click()
  'ActualizaMenus ("FBE")
  ajuste_o_cancel = 0
  frmConceptos.Show 1
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Sub ocupa_usuario()
    sql = "SELECT Nombre from Usuarios where Usuario='" & Trim(gstrUsuario) & "'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       Nombre = Mid(Trim(rsBD!Nombre), 1, 30)
       sql = "SELECT COUNT(Usuario) Total FROM usuario_ajustes"
       Set rsBD = New ADODB.Recordset
       rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
       If rsBD!total <= 0 Then
          sql = "INSERT INTO usuario_ajustes VALUES('" & Trim(gstrUsuario) & "','" & Trim(Nombre) & "')"
          cnxbdMty.Execute sql
       Else
          sql = "SELECT TOP 1 Isnull(Nombre,'USUARIO DESCONOCIDO') Usuario FROM usuario_ajustes"
          Set rsBD = New ADODB.Recordset
          rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
          Frame4.Enabled = False
          Frame3.Enabled = False
          fraAju.Enabled = False
          fraDisp.Enabled = False
          Frame2.Enabled = False
          MsgBox "¿El usuario " & Trim(rsBD!Usuario) & " esta ocupando el modulo de ajustes, y estara restringido mientras" & vbCrLf & _
          " termina sus ajustes para evitar mezclar informacion", vbInformation, "Opcion restringida Temporalmente"
       End If
    End If
End Sub

Private Sub Form_Load()
    'MsgBox gnBodega
    Set mclsAniform = New clsAnimated

    InicializaForma
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
    Call ocupa_usuario
End Sub

Private Sub Form_Terminate()
   sql = "SELECT TOP 1 * FROM usuario_ajustes"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If Not rsBD.EOF Then
      If Trim(rsBD!Usuario) = Trim(gstrUsuario) Then
         sql = "DELETE usuario_ajustes"
         cnxbdMty.Execute sql
     End If
   End If
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   sql = "SELECT TOP 1 * FROM usuario_ajustes"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If Not rsBD.EOF Then
      If Trim(rsBD!Usuario) = Trim(gstrUsuario) Then
         sql = "DELETE usuario_ajustes"
         cnxbdMty.Execute sql
     End If
   End If
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub spdAjustes_KeyPress(KeyAscii As Integer)
 With spdAjustes
    If KeyAscii = 13 Then
        If .ActiveCol = 4 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 0
            .SetFocus
            .Action = 0
            RecalculaAjustes
            
        End If
        If .ActiveCol = 1 Or .ActiveCol = 2 Then
            .Col = 1
            .Row = .ActiveRow
            Cuenta = Val(.Text)
            If Cuenta = 0 Then
                MsgBox "La cuenta no puede ser 0", vbCritical, "No ceros"
                .Col = 1
                .Text = ""
                .SetFocus
                .Action = SS_ACTION_ACTIVE_CELL
                   
            End If
            
        End If
    
    End If
End With
End Sub
Sub RecalculaAjustes()
Dim ImpAjustes As Double
ImpAjustes = 0
With spdAjustes
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        Cuenta = Val(.Text)
        If Cuenta <> 0 Then
            .Col = 4
            ImpAjustes = ImpAjustes + Val(.Text)
        End If
        
    Next i

txtImpAju = ImpAjustes
End With
End Sub
Sub RecalculaDisp()
Dim ImpDisp As Double
ImpDisp = 0
With spddisp
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        Cuenta = .Text
        If Val(Cuenta) <> 0 Then
            .Col = 4
             ImpDisp = ImpDisp + Val(.Text)
        End If
        
    Next i
txtImpDisp = ImpDisp
End With
End Sub
Private Sub spdAjustes_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
With spdAjustes
    If (Col = 1 And NewCol = 2) Or (Col = 1 And NewRow = Row + 1) Or (Col = 1 And NewRow = Row - 1) Then
        .Col = Col
        .Row = Row
        Cuenta = .Text
        If Val(Cuenta) <> 0 Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = "select empleadora, noempleado, nombre, isnull(status,1) status,saldo from cuentasbe (nolock) "
            sqls = sqls & " where nocuenta = " & Cuenta & " and Producto=" & Product
            
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
               
            If Not rsBD.EOF Then
                If rsBD!Empleadora <> Val(txtCliente.Text) Then
                   MsgBox "Esta cuenta no es de esta empresa", vbCritical, "Error en cuenta"
                   .Col = 1
                   .Text = ""
                   .SetFocus
                   .Action = SS_ACTION_ACTIVE_CELL
                                   
                Else
                    If rsBD!Status <> 2 And (cboConceptos.ItemData(cboConceptos.ListIndex) = 9 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 10) Then
                        MsgBox "Esta cuenta no esta cancelada, no puede generarse este tipo de ajuste", vbInformation, "Cuenta sin cancelar"
                         .Col = 1
                        .Text = ""
                        .SetFocus
                        .Action = SS_ACTION_ACTIVE_CELL
                   
                    Else
                        .Row = Row
                        .Col = 2
                        .Text = rsBD!noempleado
                        .Col = 3
                        .Text = rsBD!Nombre
                        .Col = 4
                        If cboConceptos.ItemData(cboConceptos.ListIndex) = 9 Or cboConceptos.ItemData(cboConceptos.ListIndex) = 10 Then
                            .Text = CDbl(rsBD!saldo)
                            ImpAj10 = CDbl(rsBD!saldo)
                            RecalculaAjustes
                        Else
                            
                            .SetFocus
                            .Action = SS_ACTION_ACTIVE_CELL
                        End If
                    End If
                   
                End If
            Else
                MsgBox "Cuenta no existe!!, favor de verificarlo", vbCritical, "Cuenta inexistente"
                .Col = 2
                .Text = ""
                .SetFocus
                .Action = SS_ACTION_ACTIVE_CELL

                Exit Sub
            End If
        End If
    
    End If 'Validacion por columnas
End With
Exit Sub
End Sub

Private Sub spdDisp_KeyPress(KeyAscii As Integer)
 With spddisp
    If KeyAscii = 13 Then
        If .ActiveCol = 4 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 0
            .SetFocus
            .Action = 0
            RecalculaDisp
            
        End If
        
        If .ActiveCol = 1 Or .ActiveCol = 2 Then
            .Col = 1
            .Row = .ActiveRow
            Cuenta = Val(.Text)
            If Cuenta = 0 Then
                MsgBox "La cuenta no puede ser 0", vbCritical, "Cuenta no debe ser cero"
                .Col = 1
                .Text = ""
                .SetFocus
                .Action = SS_ACTION_ACTIVE_CELL
            End If
        End If
    End If
End With
End Sub

Private Sub spdDisp_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
With spddisp
    If (Col = 1 And NewCol = 2) Or (Col = 1 And NewRow = Row + 1) Or (Col = 1 And NewRow = Row - 1) Then
        .Col = Col
        .Row = Row
        Cuenta = .Text
        If Val(Cuenta) <> 0 Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = "select empleadora, noempleado, nombre from cuentasbe (nolock) "
            sqls = sqls & " where nocuenta = " & Cuenta & " and Producto=" & Product
            
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
            
           
            If Not rsBD.EOF Then
                If rsBD!Empleadora <> Val(txtCliente.Text) And cboConceptos.ItemData(cboConceptos.ListIndex) <> 3 Then
                   MsgBox "Esta cuenta no es de esta empresa", vbCritical, "Error en cuenta"
                   .Col = 1
                   .Text = ""
                   .SetFocus
                   .Action = SS_ACTION_ACTIVE_CELL
                Else
                    .Row = Row
                    .Col = 2
                    .Text = rsBD!noempleado
                    .Col = 3
                    .Text = rsBD!Nombre
                    
                    .Col = 4
                    If cboConceptos.ItemData(cboConceptos.ListIndex) = 10 Then
                        .Text = ImpAj10
                        RecalculaDisp
                    Else
                        .SetFocus
                        .Action = SS_ACTION_ACTIVE_CELL
                    End If
                End If
            Else
                MsgBox "Cuenta no existe!!, favor de verificarlo", vbCritical, "Cuenta no existe"
                .Col = 2
                .Text = ""
                .SetFocus
                .Action = SS_ACTION_ACTIVE_CELL

                Exit Sub
            End If
        End If
    ElseIf Col = 4 Then
  
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .SetFocus
        .Action = 0
    
    End If 'Validacion por columnas
    
End With

End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(txtCliente.Text) <> 0 Then
        sqls = "select nombre from clientes where cliente = " & Val(txtCliente) & ""
        
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
        If Not rsBD.EOF Then
            txtNombre.Text = rsBD!Nombre
        Else
            MsgBox "Numero de cliente no existe", vbCritical, "Cliente no existe"
            txtNombre.Text = ""
            txtCliente.SetFocus
        End If
    End If
End Sub

Private Sub txtFolioCam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub txtFolioCam_LostFocus()
    
    If txtFolioCam.Text <> "" Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = "sp_AjustesVarios " & Val(txtFolioCam.Text) & "," & Product & ",'Ajuste1'"
            
'            sqls = "select folio, fechamov Fecha from ajustesbe where folio = " & Val(txtFolioCam.Text)
'            sqls = sqls & " and Producto=" & Product
            
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
        If Not rsBD.EOF Then
            mskFechaDispAnt = Format(rsBD!Fecha, "dd/mm/yyyy")
        Else
            MsgBox "Folio no existe!", vbCritical, "Folio no existe"
            txtFolioCam.Text = ""
            txtFolioCam.SetFocus
        End If
    End If

Set rsBD = Nothing
End Sub
