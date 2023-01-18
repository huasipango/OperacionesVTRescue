VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLiquidacionesBEPaso 
   Caption         =   "Integracion de  informacion de Bancos"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   120
      TabIndex        =   23
      Top             =   120
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
         ItemData        =   "frmLiquidacionesBEPaso.frx":0000
         Left            =   1800
         List            =   "frmLiquidacionesBEPaso.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label11 
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
         TabIndex        =   24
         Top             =   310
         Width           =   1545
      End
   End
   Begin VB.Frame Corrige 
      Caption         =   "Procesos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   6135
      Begin VB.CommandButton Correccion 
         BackColor       =   &H8000000D&
         Height          =   570
         Left            =   5280
         Picture         =   "frmLiquidacionesBEPaso.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Corrige, para que pueda generar nuevamente el  archivo de dispersiones"
         Top             =   600
         Width           =   570
      End
      Begin VB.CommandButton cmdProcesar 
         BackColor       =   &H8000000D&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   5280
         Picture         =   "frmLiquidacionesBEPaso.frx":0326
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   570
      End
      Begin VB.CommandButton cmdReporte 
         BackColor       =   &H8000000D&
         Enabled         =   0   'False
         Height          =   570
         Left            =   4320
         Picture         =   "frmLiquidacionesBEPaso.frx":0FF0
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   3000
         Width           =   570
      End
      Begin VB.CheckBox chkProc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chkProc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chkProc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkProc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkProc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkProc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Rechazos (RC)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   500
         TabIndex        =   21
         Top             =   3080
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Ajustes (RA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   500
         TabIndex        =   20
         Top             =   2600
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Vencimientos y Devoluciones (RV)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   500
         TabIndex        =   19
         Top             =   2100
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Dispersiones (RS)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   500
         TabIndex        =   18
         Top             =   1600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Transacciones Conc.(LQ)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   500
         TabIndex        =   17
         Top             =   1150
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Transacciones (LA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   500
         TabIndex        =   16
         Top             =   650
         Width           =   2295
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2100
         Width           =   3240
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      Begin VB.CommandButton cmdIr 
         BackColor       =   &H8000000D&
         CausesValidation=   0   'False
         Height          =   435
         Left            =   5160
         Picture         =   "frmLiquidacionesBEPaso.frx":180A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Grabar"
         Top             =   240
         Width           =   525
      End
      Begin MSMask.MaskEdBox mskFechaProc 
         Height          =   345
         Left            =   3600
         TabIndex        =   2
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAño1 
         Caption         =   "Fecha Proceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   300
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmLiquidacionesBEPaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim sqls As String
Dim prod As Byte
Dim i As Long
Dim hora_ini As Date
Sub doLoadFile(SFILE As String)
Dim nfile As Long, lngI As Long, cLine As String
Dim dTimeIni As Date, dTimeFin As Date
Dim nProducto As Integer, nAnoMes As Long, nConsecutivo As Long, _
nBodega As Integer, nValor As Double, sFecha As String, _
sFechaSurt As String
Dim sFileDir As String, nCantFile As Long
If SFILE <> "" Then
  If Mid(SFILE, Len(SFILE), 1) = "\" Then
     If Dir(SFILE, vbDirectory) <> "" Then
        dTimeIni = Now
        sFileDir = Dir(SFILE & "LQ" & Format(mskFechaProc, "yymmdd") & ".DAT", vbArchive)
        If sFileDir = "" Then
           sFileDir = Dir(SFILE & "LQ" & Format(mskFechaProc, "yymmdd"), vbArchive)
        End If
        nCantFile = 0
        lngI = 0
        If sFileDir = "" Then
            'MsgBox "No se encontró el archivo de conciliacion del dia  " & "LQ" & Format(mskFechaProc, "YYMMDD"), vbCritical, "Archivo no encontrado"
        End If
        Do While sFileDir <> ""
           nfile = FreeFile()
           Select Case Left(sFileDir, 3)
                Case "LQ1", "LQ0" 'LIQUIDACIONES
                 '   lblArchivo.Caption = "Subiendo Archivo de Liquidaciones " & sFileDir
                    DoEvents
                    resp = SubearchivoLQLA(SFILE & sFileDir, "LQ")

           End Select
           
            FileCopy SFILE & sFileDir, SFILE & "LQ\" & sFileDir
            Kill SFILE & sFileDir
            sFileDir = Dir$
        Loop
        
     Else
        MsgBox "No se encuentra la carpeta.", vbCritical, "Cree la carpeta..."
        Close #1
        Exit Sub
     End If
  End If
End If
End Sub
Private Function SubearchivoLQLA(Path As String, tipo As String) As String
Dim nArchivo, nError, clinea As String, i As Long
Dim valor As Double
Dim TiempoIni As Date, TiempoFin As Date
Dim Cuenta As Long, empleado As String
Dim noempleado As String
Dim importe As Double
Dim VALIDA As Boolean
Dim ArrVal() As String
Dim cabecera As Boolean
Dim TotReg As Long
Dim contRech As Long
Dim numTrans As Long
cabecera = False
SubearchivoLQLA = ""
valor = 0
On Error GoTo ErrorImport
nArchivo = FreeFile
Open Path For Input Access Read As #nArchivo
i = 0
contRech = 1
Screen.MousePointer = 11
Do While Not EOF(nArchivo)
   DoEvents
   Line Input #nArchivo, clinea
   Fecha = Mid(Mid(Path, InStr(1, Path, "LA") + 2, 6), 3, 2) & "/" & Mid(Mid(Path, InStr(1, Path, "LA") + 2, 6), 5, 2) & "/" & Mid(Mid(Path, InStr(1, Path, "LA") + 2, 6), 1, 2)
   If Mid(clinea, 1, 2) = "19" Then
         FechaEnvio = Mid(clinea, 16, 2) & "/" & Mid(clinea, 18, 2) & "/" & Mid(clinea, 12, 4)
         cabecera = True
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         sqls = "select id_archivo from archivosbanco where id_archivo  = '" & Trim(Mid(Path, InStrRev(Path, "\") + 1)) & "'"
         sqls = sqls & "And Producto=" & Product
          
         Set rsBD = New ADODB.Recordset
         rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   
         
         If Not rsBD.EOF And tipo = "LA" Then
            Screen.MousePointer = 1
            Progresivo = Progresivo + 1
            sqls = " EXEC sp_LoNuevo_Ins @Modulo = 'ABE', @Fecha = '" & Format(Date, "mm/dd/yyyy") & "'" & _
                   " ,@Progresivo = " & Progresivo & _
                   " ,@Texto ='El archivo del " & FechaEnvio & " no se subio , ya que ya se habian subido los movimientos de ese día'" & _
                   " ,@Version = '" & Mid(Path, InStrRev(Path, "\") + 1, 12) & "'"
            cnxBD.Execute sqls, intRegistros

            Close #nArchivo
            Exit Function
         End If
    ElseIf Mid(clinea, 1, 2) = "25" Then
         cabecera = False
         Tarjeta = Trim(Mid(clinea, 3, 19))
         fechatran = Mid(clinea, 26, 2) & "/" & Mid(clinea, 28, 2) & "/" & Mid(clinea, 22, 4) & " " & Mid(clinea, 30, 2) & ":" & Mid(clinea, 32, 2) & ":" & Mid(clinea, 34, 2)
         importe = Val(Mid(clinea, 36, 11)) / 100
         
         If Mid(clinea, 47, 1) = "-" Then
             saldo = Val(Mid(clinea, 48, 10)) / 100 * -1
         Else
             saldo = Val(Mid(clinea, 48, 10)) / 100
         End If
         
         numcomercio = Trim(Mid(clinea, 58, 9))
         secuencia = Mid(clinea, 67, 12)
         CveTran = Val(Mid(clinea, 79, 2))
         CveResp = Val(Mid(clinea, 81, 3))
         NumAut = Val(Mid(clinea, 84, 6))
         If CveResp > 1 Then 'no aprobado, con rechazo
            numTrans = contRech
            contRech = contRech + 1
         Else
            numTrans = 0
         End If
            
    ElseIf Mid(clinea, 1, 2) = "99" Then
         If Product = 6 Or Product = 8 Then
            TotReg = Val(Mid(clinea, 3, 10)) / 100
            cabecera = True
         ElseIf Product = 7 Then
            TotReg = Val(Mid(clinea, 3, 8)) / 100
            cabecera = True
         End If
    End If
     
    If cabecera = False Then
        i = i + 1
        If tipo = "LQ" Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = " exec sp_UpdLiquidacionesBE " & _
                   "  @NumAut = '" & NumAut & "'" & _
                   " ,@Tarjeta = '" & Tarjeta & "'" & _
                   " ,@FechaTran = '" & fechatran & "'" & _
                   " ,@FechaEnvio = '" & FechaEnvio & "'" & _
                   " ,@Secuencia = '" & secuencia & "'" & _
                   " ,@Producto=" & Product
        End If
            
        cnxBD.Execute sqls, intRegistros
    End If
Loop
Screen.MousePointer = 1

'sqls = "insert into archivosbanco values('" & Trim(Mid(Path, InStrRev(Path, "\") + 1)) & "'," & Product & ",getdate(), '" & FechaEnvio & "' , " & TotReg & ", 0)"
'cnxBD.Execute sqls, intRegistros
    
Progresivo = Progresivo + 1

Close #nArchivo
Exit Function
rsBD.Close
Set rsBD = Nothing

ErrorImport:
    Beep
    MsgBox "Hubo un error al actualizar! Favor de avisar a sistemas! Error: " & ERR.Number & vbCrLf & ERR.Description, vbCritical + vbOKOnly, Me.Caption
    Screen.MousePointer = 1
    Resume Next
End Function

Sub ValidaLA()
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from PasoBE where proceso = 'LA'"
sqls = sqls & " And Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    chkProc(0).value = 0
Else
    chkProc(0).value = 1
End If

End Sub

Sub ValidaLQ()
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from PasoBE where proceso = 'LQ'"
sqls = sqls & " And Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    chkProc(1).value = 0
Else
    chkProc(1).value = 1
End If

End Sub
Sub ValidaRA()

'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from PasoBE where proceso = 'RA'"
sqls = sqls & " And Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    chkProc(4).value = 0
Else
    chkProc(4).value = 1
End If


End Sub

Sub ValidaRC()

'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from PasoBE where proceso = 'RC'"
sqls = sqls & " And Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

'If Not rsBD.EOF Then
 '   chkProc(5).Value = 0
'Else
    chkProc(5).value = 1
'End If

End Sub
Sub ValidaRS()

'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from PasoBE where proceso = 'RS'"
sqls = sqls & " And Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    chkProc(2).value = 0
Else
    chkProc(2).value = 1
End If

End Sub
Sub ValidaRV()

'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from PasoBE where proceso = 'RV'"
sqls = sqls & " And Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    chkProc(3).value = 0
Else
    chkProc(3).value = 1
End If

End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  cmdIr.Enabled = True
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
     Exit Sub
  End If
End Sub

Sub valida_fecha_cierre()
  sqls = "SELECT Fecha_CierreBE FROM INFORMACION_GENERAL"
  Set rsBD = New ADODB.Recordset
  rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
  If Format(rsBD!Fecha_CierreBE, "mm/dd/yyyy") >= Format(Date - 12, "mm/dd/yyyy") Then
     sqls = "UPDATE INFORMACION_GENERAL SET Fecha_CierreBE='" & Format(Date - 20, "mm/dd/yyyy") & "'"
     cnxBD.Execute sqls
  End If
End Sub

Private Sub cmdIr_Click()
Dim statusValida As Boolean
Dim fecha_disp As String
Dim sqls As String
Screen.MousePointer = 11
statusValida = True

hora_ini = Now
fecha_disp = Format(mskFechaProc.Text, "yyyymmdd")
valida_fecha_cierre
Set rsBD = New ADODB.Recordset
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "EXEC Sp_ValidaLiquidacionesBE  '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "', 'RE'," & Product
cnxBD.Execute sqls, intRegistros

ValidaLA
ValidaLQ
ValidaRS
ValidaRV
ValidaRA
ValidaRC

For i = 0 To 5
    If chkProc(i).value = False Then
        statusValida = False
    End If
Next i
    
If statusValida = False Then
    MsgBox "No se pueden integrar la informacion del día, ya que contiene errores", vbCritical, "Errores encontrados"
    cmdProcesar.Enabled = False
    Correccion.Visible = True
    Call cmdReporte_Click
Else
    MsgBox "La informacion es correcta , ya puede integrarlo a la consulta de Saldos", vbInformation, "Informacion correcta"
    cmdProcesar.Enabled = True
End If
Screen.MousePointer = 1
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptValidacionLiqBEPaso.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Error..."
    End If
    
End Sub

Private Sub cmdProcesar_Click()
Dim hora_fin As Date
 
 Screen.MousePointer = 11
 'prod = IIf(Product = 8, 6, Product)
 producto_cual
 sqls = "exec sp_PaseLiquidacionesBE '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & Product
 cnxBD.Execute sqls, intRegistros
 
 'Call doLoadFile("C:\Bono Electronico\ArchivosSistema\Resp\")
 'If Product <> 9 And Product <> 10 And Product <> 11 Then
 sqls = "SELECT * FROM archivos_automaticos Where Producto=" & Product & " and Fecha_proceso='" & Format(mskFechaProc, "mm/dd/yyyy") & "'"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 If Not rsBD.EOF Then
    hora_fin = Now
    sqls = "UPDATE archivos_automaticos SET Hora_termino='" & Format(CDate(hora_fin), "yyyy-mm-dd HH:MM:SS") & "',"
    sqls = sqls & " Status=4"
    sqls = sqls & " Where Fecha_proceso='" & Format(mskFechaProc, "mm/dd/yyyy") & "' AND Producto=" & Product
    cnxBD.Execute sqls
 Else
    hora_fin = Now
    sqls = "INSERT INTO archivos_automaticos Values('" & Format(mskFechaProc, "mm/dd/yyyy") & "',"
    sqls = sqls & Product & ",4,'Todos','" & Format(CDate(hora_ini), "yyyy-mm-dd HH:MM:SS") & "',"
    sqls = sqls & "'" & Format(CDate(hora_fin), "yyyy-mm-dd HH:MM:SS") & "')"
    cnxBD.Execute sqls
 End If
 'End If
 
 MsgBox "La integracion del día " & mskFechaProc.Text & "  se ha realizado", vbInformation, "Integracion correcta"
 Screen.MousePointer = 1
 'Unload Me
 InicializaForma
 
End Sub

Private Sub cmdReporte_Click()
    Imprime crptToWindow
End Sub

Private Sub Correccion_Click()
On Error GoTo error_gral:
    Dim fecha_disp As String
    Dim sqls As String
    fecha_disp = Format(mskFechaProc.Text, "yyyymmdd")
    Set rsBD = New ADODB.Recordset
    
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "EXEC sp_corrige_dispersiones '" & fecha_disp & "'," & Product
    cnxBD.Execute sqls
    MsgBox "Se ha corregido el error en dispersiones" & vbCrLf & "Vuelva a generarlo", vbInformation, "Proceso completado exitosamente"
    Correccion.Visible = False
    Set rsBD = Nothing
    Exit Sub
error_gral:
   MsgBox "Error ==> " & ERR.Description, vbCritical, "No se puedo realizar el proceso"
   Exit Sub
End Sub


Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
On Error GoTo ERRO:
    Set mclsAniform = New clsAnimated
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
    InicializaForma
Exit Sub
ERRO:
  MsgBox "Se han encontrado errores..." & ERR.Description, vbCritical, "Imposible continuar..."
  Exit Sub
End Sub

Sub InicializaForma()
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "select ISNULL(max(fechaenvio),'01/01/1900')Fecha from liquidacionesbepaso" & _
           " where status < 2 and comercio not in ('AJUS')" 'and cveresp in (0,1)
    sqls = sqls & " AND Producto=" & Product

    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Correccion.Visible = False
    Call limpia_chek
    If rsBD!Fecha <> "01/01/1900" Then
        mskFechaProc.Text = Format(rsBD!Fecha, "dd/mm/yyyy")
    Else
        MsgBox "No hay transacciones pendientes", vbInformation, "Sistema Bono Electronico"
        cmdIr.Enabled = False
        Exit Sub
    End If
    cmdProcesar.Enabled = False
End Sub

Function corrige_fecha(fec As String) As String
Dim dia, mes, ano As String
dia = Mid(Fecha, 1, 2)
mes = Mid(Fecha, 4, 2)
ano = Mid(Fecha, 7, 4)
corrige_fecha = ano & mes & dia
End Function

Sub limpia_chek()
Dim X As Byte
For X = 0 To 5
    chkProc(X).value = 0
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

