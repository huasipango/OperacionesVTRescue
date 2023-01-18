VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepAclara 
   Caption         =   "Aclaraciones"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo (dd/mm/yyyy) "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   960
         TabIndex        =   7
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   345
         Left            =   960
         TabIndex        =   8
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   915
         Width           =   255
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   735
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   120
         Picture         =   "frmRepAclara.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   120
         Picture         =   "frmRepAclara.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cbostatus 
         Height          =   315
         ItemData        =   "frmRepAclara.frx":0274
         Left            =   1200
         List            =   "frmRepAclara.frx":0276
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFolio 
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
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRepAclara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPresentar_Click()
    Imprime crptToWindow
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Product = 1 Then
       Frame2.Caption = "Winko Mart"
    Else
       Frame2.Caption = "Vale Total Combustible"
    End If
    CargaStatus cbostatus
    cbostatus.AddItem "TODAS"
    cbostatus.ItemData(2) = 0
       
    Call CboPosiciona(cbostatus, 0)
    txtFolio.Text = 0
    mskFechaIni.Text = "01/01/" & Year(Date)
    mskFechaFin.Text = Date
    
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAclaraciones.rpt"
    'mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFacoi2.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(txtFolio.Text)
    mdiMain.cryReport.StoredProcParam(1) = CStr(cbostatus.ItemData(cbostatus.ListIndex))
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(4) = CStr(Product)
    
    'mdiMain.cryReport.Formulas(0) = mskFechaIni
    'mdiMain.cryReport.Formulas(1) = mskFechaFin
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub
