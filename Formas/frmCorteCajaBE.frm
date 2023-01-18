VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCorteCajaBE 
   Caption         =   "Corte de Caja Proveedores BE"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorteCajaBE.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorteCajaBE.frx":4889A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboBodegas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   3615
   End
   Begin MSMask.MaskEdBox mskFechaFin 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Enc"
      ToolTipText     =   "Fecha del Movimiento"
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1852
      ButtonWidth     =   1455
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar"
            Key             =   "Mostrar"
            Object.ToolTipText     =   "Muestra estado de cuenta"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Corte:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2475
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   810
   End
End
Attribute VB_Name = "frmCorteCajaBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated, Iva As Double
Dim prod As Byte


Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  CargaBodegas2 cboBodegas
  mskFechaFin = Format(Date, "dd/mm/yyyy")
  If tipo_estad = 1 Then
     Me.Caption = "Corte de Caja (Otros Ingresos)"
  Else
     Me.Caption = "Aplicacion de la cobranza (Otros Ingresos)"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Mostrar"
         Call Mostrar
End Select
End Sub

Sub Mostrar()
Dim Result As Integer
    If tipo_estad = 1 Then
       sqls = "sp_Prepara_CorteCajaOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & ",'"
       sqls = sqls & Format(mskFechaFin, "mm/dd/yyyy") & "','Prepara'"
       cnxBD.Execute sqls
       
       MsgBar "Generando Reporte", True
       Limpia_CryReport
       mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
       mdiMain.cryReport.ReportFileName = gPath & "\Reportes\Corte_CajaOI.rpt" 'NO ES NECESARIO CAMBIARLO
       mdiMain.cryReport.Destination = Destino
       mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
       mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
       mdiMain.cryReport.StoredProcParam(2) = 0
        
       On Error Resume Next
       Result = mdiMain.cryReport.PrintReport
       MsgBar "", False
       If Result <> 0 Then
          MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
       End If
    End If
    If tipo_estad = 2 Then
       sqls = "sp_Prepara_CorteCajaOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & ",'"
       sqls = sqls & Format(mskFechaFin, "mm/dd/yyyy") & "','Prepara'"
       cnxBD.Execute sqls
       
       MsgBar "Generando Reporte", True
       Limpia_CryReport
       mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
       mdiMain.cryReport.ReportFileName = gPath & "\Reportes\apl_cob.rpt" 'NO ES NECESARIO CAMBIARLO
       mdiMain.cryReport.Destination = Destino
       mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
       mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
       mdiMain.cryReport.StoredProcParam(2) = 0
        
       On Error Resume Next
       Result = mdiMain.cryReport.PrintReport
       MsgBar "", False
       If Result <> 0 Then
          MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
       End If
    End If
End Sub

