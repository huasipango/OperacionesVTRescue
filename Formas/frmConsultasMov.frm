VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsultasMov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Otros Ingresos"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   3960
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   570
         Left            =   120
         Picture         =   "frmConsultasMov.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   1320
         Width           =   1170
      End
      Begin VB.CommandButton cmdPresentar 
         Caption         =   "Imprimir"
         Height          =   570
         Left            =   120
         Picture         =   "frmConsultasMov.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   480
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3735
      Begin VB.ListBox lstMovClientes 
         Height          =   2010
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox cboSucursal 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Key"
         Top             =   240
         Width           =   3840
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   3960
         TabIndex        =   5
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Fecha Final:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label lblBodega 
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmConsultasMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated

Private Sub CargaTiposMovimientos()
   
   sqls = " SELECT Tipo_mov, Descripcion "
   sqls = sqls & vbCr & "  FROM fm_Tipos_mov_cartera "
    
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   
   i = 0
   Do While Not rsBD.EOF
      lstMovClientes.AddItem rsBD!descripcion
      lstMovClientes.ItemData(i) = rsBD!TIPO_MOV
      i = i + 1
      rsBD.MoveNext
   Loop
    
    If lstMovClientes.ListCount > 0 Then
       lstMovClientes.ListIndex = 0
    Else
      MsgBox "No existen Movimientos dados de Alta", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdPresentar_Click()
    Imprime crptToWindow
End Sub

Private Sub cmdSalir_Click()
        Unload Me
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    
    CargaBodegas CboSucursal
    
    mskFechaIni = Format((Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Trim(Str(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))))), "mm/dd/yyyy")
    mskFechaFin = Format((FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date) - 1), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)))), "mm/dd/yyyy")
    Call CargaTiposMovimientos

End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    
    strTipoPro = ""
   
    For i = 0 To lstMovClientes.ListCount - 1
      If lstMovClientes.Selected(i) = True Then
         If Len(strTipoPro) >= 1 Then
            strTipoPro = strTipoPro & "," & lstMovClientes.ItemData(i)
         Else
            strTipoPro = strTipoPro & lstMovClientes.ItemData(i)
         End If
      End If
    Next i

    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFacoi.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = CboSucursal.ItemData(CboSucursal.ListIndex)
    mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = strTipoPro
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
