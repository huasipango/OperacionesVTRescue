VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmCancCuentas 
   Caption         =   "Baja de Tarjetas Titulares y Adicionales"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   480
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
         ItemData        =   "frmCancCuentas.frx":0000
         Left            =   1800
         List            =   "frmCancCuentas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
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
         TabIndex        =   10
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   10575
      Begin FPSpread.vaSpread spddetalle 
         Height          =   2535
         Left            =   120
         OleObjectBlob   =   "frmCancCuentas.frx":001C
         TabIndex        =   3
         Top             =   240
         Width           =   10245
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10575
      Begin VB.CommandButton cmdAbrircsv 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   1440
         Picture         =   "frmCancCuentas.frx":033A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Subir desde archivo"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   240
         Picture         =   "frmCancCuentas.frx":0954
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdBorrar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   2040
         Picture         =   "frmCancCuentas.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borra registro"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   840
         Picture         =   "frmCancCuentas.frx":0B58
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Subir desde archivo"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   2640
         Picture         =   "frmCancCuentas.frx":0C5A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cancelación de Tarjetas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7560
         TabIndex        =   8
         Top             =   285
         Width           =   2910
      End
   End
   Begin MSComDlg.CommonDialog cmnAbrir 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCancCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte

Private Sub cmdAbrir_Click()
   Call Subearchivo
End Sub

Private Function Subearchivo()

Dim nArchivo, nError, clinea As String, i As Long
Dim valor As Double
Dim TiempoIni As Date, TiempoFin As Date
Dim Cuenta As Long, empleado As String
Dim noempleado As String, tipocan As Byte
Dim importe As Double
Dim VALIDA As Boolean
Dim ArrVal() As String
Dim cabecera As Boolean
nArchivo = "c:\Facturacion\LogError.txt"
Close #1
Open nArchivo For Output As #1

With spdDetalle
    Subearchivo = ""
    On Error GoTo ErrorImport
    cmnAbrir.ShowOpen
    If cmnAbrir.Filename <> "" Then
        nArchivo = FreeFile
        Open cmnAbrir.Filename For Input Access Read As #nArchivo
        'Open cmnAbrir.Filename For Random As #nArchivo
        i = 0
    
        Do While Not EOF(nArchivo)
            DoEvents
            
            Line Input #nArchivo, clinea
            If Mid(clinea, 1, 2) = "07" Then
                 lblEmp = Val(Mid(clinea, 40, 7))
                 cabecera = True
            ElseIf Mid(clinea, 1, 2) = "08" Then
                 i = i + 1
                 .Row = i
                 .MaxRows = i
                 cliente = Val(Mid(clinea, 3, 5))
                 Cuenta = Mid(clinea, 8, 8)
                 Tarjeta = Mid(clinea, 23, 16)
                 nombre = Trim(Mid(clinea, 39, 26))
                 tipocan = Trim(Mid(clinea, 65, 1))
                 'prod = IIf(Product = 8, 6, Product)
                 producto_cual
                 sqls = "sp_CuentasBE_Varios " & Cuenta & "," & Product & ",'Cuenta'"
                 
'                 sqls = "select * from cuentasbe where nocuenta = " & cuenta
'                 sqls = sqls & " and Producto=" & Product
                 
                 Set rsBD = New ADODB.Recordset
                 rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                 obs = ""
                 If rsBD.EOF Then
                    obs = "No existe la cuenta"
                 Else
                    If rsBD!Empleadora <> cliente Then
                        obs = "La cuenta no corresponde a ese cliente"
                    End If
                    
                    If rsBD!Status = 2 Then
                        obs = "Esta cuenta ya esta cancelada"
                    End If
                    
                    If rsBD!saldo > 0 Then
                        obs = "Cuenta con Saldo"
                    End If
                    
                 End If
                 .Col = 1
                 .Text = cliente
                 .Col = 2
                 .Text = Cuenta
                 .Col = 3
                 .Text = Tarjeta
                 .Col = 4
                 .Text = nombre
                 .Col = 5
                  If Val(tipocan) = 0 Or Val(tipocan) = 3 Then
                     .Text = "Inactividad de Tarjeta"
                  ElseIf Val(tipocan) = 1 Then
                     .Text = "Nota de Credito"
                  ElseIf Val(tipocan) = 2 Then
                     .Text = "Baja de Cliente"
                  End If
                
                 cabecera = False
            End If
          
            
        Loop
        
        MsgBar "Listo...", False
        
        Screen.MousePointer = 1
        
        lblEmp = i & " Cuentas por cancelar"
        
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

Private Sub cmdAbrircsv_Click()
Dim nArchivo, clinea As String, i As Long
Dim valor As Double
Dim TiempoIni As Date, TiempoFin As Date
Dim Cuenta As Long, empleado As Long
Dim ImpVentas, impdev, ImpTotal As Double
Dim TotInter, TotDeb, TotPos As Double, tipocan As Byte
Dim NoInter, NoDeb, NoPos, tipocom As Integer
Dim nombre As String
Dim ArrVal() As String
Dim ArrImp() As Double
Dim ArrCant() As Integer
Screen.MousePointer = 11
NumEmp = 0
With spdDetalle
    .Col = -1
    .Row = -1
    .Action = 12
    cmnAbrir.ShowOpen
    Screen.MousePointer = 1
    If cmnAbrir.Filename <> "" Then
        nArchivo = FreeFile
        Open cmnAbrir.Filename For Input Access Read As #nArchivo
        i = 0
        
        On Error GoTo err_file
        Do While Not EOF(nArchivo)
            Line Input #nArchivo, clinea
            ArrVal = Split(clinea, ",")
            
            cliente = CLng(ArrVal(0))
           ' If Not ValidaExiste(CLng(ArrVal(0)), QuitaCeros(ArrVal(3)), IIf(ArrVal(2) = 0, "T", "A")) Then
           '     If ValidaExisteTitular(CLng(ArrVal(0)), QuitaCeros(ArrVal(3)), IIf(ArrVal(2) = 0, "T", "A")) Then
                
                     i = i + 1
                     .Row = i
                     .MaxRows = i
                     .Col = 1
                     .Text = ArrVal(0)
                     cliente = CLng(ArrVal(0))
                     'nombre = BuscaCliente(Cliente, nombre)
                     .Col = 2
                     .Text = ArrVal(1) 'Trim(nombre)
                     .Col = 3
                     .Text = ArrVal(2)
                     .Col = 4
                     .Text = ArrVal(3)
                     .Col = 5
                     tipocan = ArrVal(4)
                     If Val(tipocan) = 0 Or Val(tipocan) = 3 Then
                        .Text = "Inactividad de Tarjeta"
                     ElseIf Val(tipocan) = 1 Then
                        .Text = "Nota de Credito"
                     ElseIf Val(tipocan) = 2 Then
                        .Text = "Baja de Cliente"
                     End If
                                          
                     DoEvents
            '    End If
            'End If
        Loop
        Close #nArchivo
        Screen.MousePointer = 1

        cmnAbrir.Filename = ""
        Exit Sub
    Else
        Exit Sub
    End If
End With
Screen.MousePointer = 1
Exit Sub
err_file:
    MsgBox "Error en el formato del archivo" & vbCrLf & "Quizas le falta agregar el motivo de cancelacion", vbCritical, "Archivo incorrecto"
    Screen.MousePointer = 1
End Sub

Private Sub cmdBorrar_Click()
With spdDetalle
    .Row = .ActiveRow
    .Col = -1
    .Action = 5
    .MaxRows = .MaxRows - 1
End With
End Sub

Private Sub cmdGrabar_Click()
Dim tipocan As Byte
 If spdDetalle.MaxRows = 0 Then Exit Sub
 If MsgBox("Esta seguro de que desea cancelar estas cuentas?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancelacion de Cuentas") = vbYes Then
        With spdDetalle
            For i = 1 To .MaxRows
                .Row = i
                .Col = 2
                Cuenta = .Text
                .Col = 3
                Tarjeta = .Text
                .Col = 5
                If Trim(.Text) = "" Then
                   tipocan = 3
                ElseIf Trim(.Text) = "Nota de Credito" Then
                   tipocan = 1
                ElseIf Trim(.Text) = "Baja de Cliente" Then
                   tipocan = 2
                ElseIf Trim(.Text) = "Inactividad de Tarjeta" Then
                   tipocan = 3
                End If
                'prod = IIf(Product = 8, 6, Product)
                producto_cual
                sqls = "sp_CuentasBE_Varios 0," & Product & ",'Tarjeta','" & Tarjeta & "'"
                
'                sqls = "select tipo from tarjetasbe where notarjeta = '" & TARJETA & "'"
'                sqls = sqls & " and Producto=" & Product
                
                Set rsBD = New ADODB.Recordset
                rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                
                If Not rsBD.EOF Then
                    If Trim(rsBD!tipo) = "T" Or Trim(rsBD!tipo) = "RT" Then
                        'prod = IIf(Product = 8, 6, Product)
                        producto_cual
                        sqls = "sp_CuentasBE_Varios " & Cuenta & "," & Product & ",'Cancela'"
                        
'                        sqls = "update cuentasbe set  status = 2" & _
'                               " where nocuenta = " & cuenta & _
'                               " and Producto=" & Product
                        
                        cnxbdMty.Execute sqls, intRegistros
                    End If
                End If
                'prod = IIf(Product = 8, 6, Product)
                producto_cual
                sqls = "sp_CuentasBE_Varios " & tipocan & "," & Product & ",'CancelaTar','" & Tarjeta & "'"
                
'                sqls = "update tarjetasbe set  status = 2, fechacancelacion = getdate()" & _
'                       " where notarjeta  = '" & TARJETA & "' and Producto=" & Product
                
                cnxbdMty.Execute sqls, intRegistros
               
            Next i
        .Col = -1
        .Row = -1
        .Action = 12
        .MaxRows = 0
      
        End With
        
        MsgBox "La cuentas han sido canceladas!!", vbInformation, "Cuentas canceladas"
    End If
    
'GeneraArchivo
    
End Sub
Sub GeneraArchivo()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, nombre
Dim Empleadora As String
Dim RetVal

On Error GoTo err_gral

Close #1

Open "c:\Facturacion\Cancelacion" & Format(Date, "DDMM") & ".txt" For Output As #1

clinea = "071003602Vale Total S.A. DE C.V.       "
clinea = clinea & Format(spdDetalle.MaxRows, "0000000") & Format(Date, "DD-MMM-YY")
Print #1, clinea
               
With spdDetalle

 For i = 1 To .MaxRows
       clinea = "08"
       .Row = i
       .Col = 1
       cliente = Val(.Text)
       clinea = clinea & Format(Val(.Text), "00000") '& Space(5) & Format(i, "0000000")
       .Col = 2
       Cuenta = .Text
       clinea = clinea & Format(Val(Cuenta), "00000000")
       clinea = clinea & Format(i, "0000000")
       .Col = 3
       clinea = clinea & .Text
       .Col = 4
       clinea = clinea & Pad(Trim(.Text), 26, " ", "R")
       Print #1, clinea
 Next i
 
 Close #1
End With

RetVal = Shell("C:\WINDOWS\NOTEPAD.EXE c:\Facturacion\Cancelacion" & Format(Date, "DDMM") & ".txt", 1)
Exit Sub

err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Solicitud de Tarjetas"
   Call doErrorLog(gnBodega, "FACBE", ERR.Number, ERR.Description, Usuario, "frmCargaSucursales.GeneraArchivo")
   MsgBar "", False
   Exit Sub
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
  'Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
End Sub

Sub InicializaForma()
    spdDetalle.MaxRows = 1
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub
Private Sub Form_Load()
Set mclsAniform = New clsAnimated

cboProducto.Clear
Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
cboProducto.Text = UCase("Despensa Total")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
