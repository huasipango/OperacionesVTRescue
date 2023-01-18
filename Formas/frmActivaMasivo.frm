VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmActivaMasivo 
   Caption         =   "Activacion de Cuentas de Stock de Manera Masiva "
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   550
      Left            =   240
      TabIndex        =   4
      Top             =   840
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
         ItemData        =   "frmActivaMasivo.frx":0000
         Left            =   1800
         List            =   "frmActivaMasivo.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   140
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
         TabIndex        =   6
         Top             =   180
         Width           =   1545
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   8640
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proporcione los datos de acuerdo al orden mencionado"
      ForeColor       =   &H00808080&
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11415
      Begin FPSpread.vaSpread spddetalle 
         Height          =   5895
         Left            =   240
         OleObjectBlob   =   "frmActivaMasivo.frx":001C
         TabIndex        =   2
         Top             =   360
         Width           =   10935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1111
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivo"
            Key             =   "Archivo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            Key             =   "Guardar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivaMasivo.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivaMasivo.frx":0A46
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivaMasivo.frx":492E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Formato del archivo: [ No.Cliente ] , [ Cuenta  ] , [ No. Empleado ] , [ Nombre Empleado ] , [Plaza ]"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   8280
      Width           =   6870
   End
End
Attribute VB_Name = "frmActivaMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
   Set mclsAniform = New clsAnimated
   CboProducto.Clear
   Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
   CboProducto.Text = UCase("Winko Mart")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Archivo"
         abrir
   Case "Guardar"
         Grabar
   End Select
End Sub

Sub abrir()
Dim ArrVal() As String
Dim nArchivo As Byte, clinea As String, i As Long, Tarjeta As String, tipo As String
Dim Nombre As String, empleado As String, NumEmp As Integer, Cuenta As Long
Dim ncampos As Byte, cliente As Long, nomcte As String, Plaza As String, banda_activa As Boolean
On Error GoTo ERR:
With spddetalle
    .Col = -1
    .Row = -1
    .Action = 12
    cmd.ShowOpen
    If cmd.Filename <> "" Then
       nArchivo = FreeFile
       Open cmd.Filename For Input Access Read As #nArchivo
       i = 0
       NumEmp = 0
       Do While Not EOF(nArchivo)
          Line Input #nArchivo, clinea
          ArrVal = Split(clinea, ",")
          ncampos = UBound(ArrVal)
          If ncampos < 4 Then
             Exit Do
          End If
          cliente = CLng(ArrVal(0))
          Cuenta = CLng(ArrVal(1))
          empleado = Trim(UCase(ArrVal(2)))
          empleado = QuitaCeros(empleado)
          Nombre = Trim(UCase(ArrVal(3)))
          Nombre = Trim(Mid(Nombre, 1, 26))
          Plaza = Trim(UCase(ArrVal(4)))
          banda_activa = False
          If Not ValidaExiste(cliente, QuitaCeros(empleado), "T", Cuenta) Then
                 i = i + 1
                 .Row = i
                 .MaxRows = i
                 .Col = 1
                 .Text = cliente
                  nomcte = BuscaCliente(cliente, nomcte)
                 .Col = 2
                 .Text = Trim(nomcte)
                 .Col = 3
                 .Text = Val(Cuenta)
                 .Col = 4
                 .Text = Trim(empleado)
                 .Col = 5
                 .Text = Trim(Nombre)
                  If Not ValidaNombre(Trim(Nombre)) Then
                     .Col = 5
                     .ForeColor = RGB(255, 0, 0)
                     NumEmp = NumEmp + 1
                  Else
                     .Col = 5
                     .ForeColor = RGB(0, 0, 0)
                  End If
                  .Col = 6
                  .Text = Trim(Plaza)
                  DoEvents
          Else
                 banda_activa = True
                 i = i + 1
                 .Row = i
                 .MaxRows = i
                 .Col = 1
                 .Text = "XX"
                 .Col = 2
                 .ForeColor = vbRed
                 .Text = "CUENTA YA ESTA ACTIVA"
                 .Col = 3
                 .ForeColor = vbRed
                 .Text = Val(Cuenta)
                 .Col = 4
                 .Text = Trim(empleado)
                 .Col = 5
                 .Text = Trim(Nombre)
                 .Col = 6
                 .Text = "XX"
                  DoEvents
          End If
          '******valida si existe algun empleado con ese cliente
          If banda_activa = False Then
             If Not ValidaExisteEmp(cliente, QuitaCeros(empleado), "T", Cuenta) Then
          
             Else
                 .Row = i
                 .MaxRows = i
                 .Col = 1
                 .Text = "XX"
                 .Col = 2
                 .Text = "EMPLEADO YA TIENE CUENTA"
                 .ForeColor = vbRed
                 .Col = 3
                 .ForeColor = vbRed
                 .Text = Val(Cuenta)
                 .Col = 4
                 .Text = Trim(empleado)
                 .Col = 5
                 .Text = Trim(Nombre)
                 .Col = 6
                 .Text = "XX"
                  DoEvents
             End If
          End If
       Loop
       Close #nArchivo
       cmd.Filename = ""
    End If
End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Function ValidaExiste(cliente As Long, empleado As String, tipo As String, noCuenta As Long)
Dim rsexist As ADODB.Recordset, rsexist2 As ADODB.Recordset

sqls = "sp_SolicitudesBE_varios '" & Trim(cliente) & "','" & tipo & "','" & Trim(empleado) & "'," & Product & ",'Filtro1',Null," & Val(noCuenta)
Set rsexist = New ADODB.Recordset
rsexist.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

If rsexist.EOF Then
   ValidaExiste = False
   Exit Function
Else
   ' MsgBox "Este numero de cuenta ya esta asignado a otro empleado en este producto", vbCritical, "Cuenta ya existe [" & Val(noCuenta) & "]"
End If

ValidaExiste = True
End Function

Function ValidaExisteEmp(cliente As Long, empleado As String, tipo As String, noCuenta As Long)
Dim rsexist As ADODB.Recordset, rsexist2 As ADODB.Recordset

sqls = "sp_SolicitudesBE_varios '" & Trim(cliente) & "','" & tipo & "','" & Trim(empleado) & "'," & Product & ",'Filtro3',Null," & Val(noCuenta)
Set rsexist = New ADODB.Recordset
rsexist.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

If rsexist.EOF Then
   ValidaExisteEmp = False
   Exit Function
Else
   ' MsgBox "Este numero de cuenta ya esta asignado a otro empleado en este producto", vbCritical, "Cuenta ya existe [" & Val(noCuenta) & "]"
End If

ValidaExisteEmp = True
End Function

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Function ValidaNombre(Nombre As String)
If Len(Trim(Nombre)) > 26 Then
    ValidaNombre = False
    Exit Function
End If

For j = 1 To Len(Nombre)
     If (Asc(Mid(Nombre, j, 1)) <= 47 Or Asc(Mid(Nombre, j, 1)) >= 91) And Asc(Mid(Nombre, j, 1)) <> 32 Then
        ValidaNombre = False
        Exit Function
    End If
Next j

For j = 1 To Len(Nombre)
     If (Asc(Mid(Nombre, j, 1)) >= 58 And Asc(Mid(Nombre, j, 1)) <= 64) Then
        ValidaNombre = False
        Exit Function
    End If
Next j

ValidaNombre = True
End Function

Sub Grabar()
Dim cliente As Long, empleado As String, Cuenta As Long, Tarjeta As String, tipo As String
Dim i As Integer, conta As Integer, Nombre As String, Plaza As String
On Error GoTo ERR
tipo = "T"
If MsgBox("¿Esta absolutamente seguro de activar masivamente todo el contenido de cuentas que acaba de cargar?", vbQuestion + vbDefaultButton2 + vbYesNo, "A punto de grabar...") = vbYes Then
   With spddetalle
      conta = 0
      For i = 1 To .MaxRows
          .Row = i
          .Col = 1
          cliente = Val(.Text)
          .Col = 3
          Cuenta = Val(.Text)
          .Col = 4
          empleado = Trim(UCase(.Text))
          .Col = 5
          Nombre = Mid(UCase(Trim(.Text)), 1, 26)
          .Col = 6
          Plaza = Trim(.Text)
          If Plaza <> "XX" Then
             sqls = "sp_Activacion_Tarjeta " & Product & "," & Val(cliente) & ",NULL," & Val(Cuenta) & ",'" & UCase(empleado) & "','" & UCase(Trim(Nombre)) & "'," & Plaza & ",'Activa2'"
             cnxbdMty.Execute sqls
             conta = conta + 1
          End If
      Next
      MsgBox "Se han activado " & conta & " cuentas de Stock", vbInformation, "Terminó proceso de activación masiva"
      .MaxRows = 0
   End With
End If
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub


