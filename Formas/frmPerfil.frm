VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmPerfil 
   Caption         =   "Usuarios ERP Vale Total"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spdDetalle 
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "frmPerfil.frx":0000
      TabIndex        =   1
      Top             =   1200
      Width           =   8175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerfil.frx":038B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerfil.frx":48C25
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerfil.frx":914BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPerfil.frx":D9D59
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1852
      ButtonWidth     =   1455
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar usuario master"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inactivar"
            Key             =   "Inactivar"
            Object.ToolTipText     =   "Inactivar usuario"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Graba informacion"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Call Cargar
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Agregar"
         Agregar
   Case "Inactivar"
         Inactivar
   Case "Grabar"
         Grabar
   End Select
End Sub
Sub Cargar()
Dim sqls As String, i As Integer

On Error GoTo ERR:
   sqls = "sp_Empleados_Sel"
   Set rsBD2 = New ADODB.Recordset
   rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   With spddetalle
       .Col = -1
       .Row = -1
       .Action = 12
       .MaxRows = 0
       i = 1
       Do While Not rsBD2.EOF
          .MaxRows = i
          .Row = i
          .Col = 1
          .Text = rsBD2!empleado
          .Col = 2
          .Text = UCase(rsBD2!Nombre)
          .Col = 3
          .Lock = False
          If rsBD2!Puesto = 1 Then
            .Text = "Gerencia de Desarrollo e Inov."
          ElseIf rsBD2!Puesto = 2 Then
            .Text = "Tesoreria"
          ElseIf rsBD2!Puesto = 3 Then
            .Text = "Supervisor Ventas"
          ElseIf rsBD2!Puesto = 4 Then
            .Text = "Gerencia de Operaciones"
          ElseIf rsBD2!Puesto = 5 Then
            .Text = "Coord. de Sistemas e Infraest."
          ElseIf rsBD2!Puesto = 6 Then
            .Text = "Ejecutivo Ventas"
          ElseIf rsBD2!Puesto = 7 Then
            .Text = "Ejecutivo Operaciones"
          ElseIf rsBD2!Puesto = 8 Then
            .Text = "Ejecutivo Atencion a Clientes"
          ElseIf rsBD2!Puesto = 9 Then
            .Text = "Director de Ventas"
          ElseIf rsBD2!Puesto = 10 Then
            .Text = "Contabilidad"
          ElseIf rsBD2!Puesto = 11 Then
            .Text = "Ventas y Atencion a Clientes"
          Else
            .Text = ""
          End If
          DoEvents
          i = i + 1
          rsBD2.MoveNext
       Loop
  End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub Inactivar()
Dim numero As String, Nombre As String, Puesto As Integer, DescPuesto As String
On Error GoTo err_gral
With spddetalle
 If .MaxRows > 0 Then
      .Row = .ActiveRow
      .Col = 1
      numero = .Text
      .Col = 2
      Nombre = UCase(.Text)
      .Col = 3
      DescPuesto = .Text
      If DescPuesto = "Gerencia de Desarrollo e Inov." Then
            Puesto = 1
      ElseIf DescPuesto = "Tesoreria" Then
            Puesto = 2
      ElseIf DescPuesto = "Supervisor Ventas" Then
            Puesto = 3
      ElseIf DescPuesto = "Gerencia de Operaciones" Then
            Puesto = 4
      ElseIf DescPuesto = "Coord. de Sistemas e Infraest." Then
            Puesto = 5
      ElseIf DescPuesto = "Ejecutivo Ventas" Then
            Puesto = 6
      ElseIf DescPuesto = "Ejecutivo Operaciones" Then
            Puesto = 7
      ElseIf DescPuesto = "Ejecutivo Atencion a Clientes" Then
            Puesto = 8
      ElseIf DescPuesto = "Director de Ventas" Then
            Puesto = 9
      ElseIf DescPuesto = "Contabilidad" Then
            Puesto = 10
      ElseIf DescPuesto = "Ventas y Atencion a Clientes" Then
            Puesto = 11
      End If
      
      If MsgBox("¿Esta seguro de que desea inactivar el empleado " & Nombre & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Inactivando usuario") = vbYes Then
         
         sqls = "sp_Empleados_Upd  " & numero & ",'" & Nombre & "'," & Puesto & ", 0, '" & DescPuesto & "'"
         cnxbdMty.Execute sqls
         
         .Row = .ActiveRow
         .Action = 5
         .MaxRows = .MaxRows - 1
      End If
      Cargar
 End If
End With
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Errores generados"
   Exit Sub
End Sub

Sub Agregar()
    With spddetalle
      .Col = 1
      .Row = .MaxRows + 1
      .MaxRows = .MaxRows + 1
      .Col = 1
      .Row = .MaxRows
      .Col = 3
      .Lock = False
      .Text = "Gerencia de Desarrollo e Inov."
      .Col = 1
      .Action = 0
      .SetFocus
    End With
End Sub

Sub Grabar()
Dim sqls As String, i As Integer
Dim numero As String, Nombre As String, Puesto As Integer, DescPuesto As String
On Error GoTo ERR:
With spddetalle
     .Row = .MaxRows
     .Col = 1
        For i = 1 To .MaxRows
             .Row = i
             .Col = 1
             numero = .Text
             .Col = 2
             Nombre = UCase(.Text)
             .Col = 3
             DescPuesto = .Text
             If DescPuesto = "Gerencia de Desarrollo e Inov." Then
                   Puesto = 1
             ElseIf DescPuesto = "Tesoreria" Then
                   Puesto = 2
             ElseIf DescPuesto = "Supervisor Ventas" Then
                   Puesto = 3
             ElseIf DescPuesto = "Gerencia de Operaciones" Then
                   Puesto = 4
             ElseIf DescPuesto = "Coord. de Sistemas e Infraest." Then
                   Puesto = 5
            ElseIf DescPuesto = "Ejecutivo Ventas" Then
                  Puesto = 6
            ElseIf DescPuesto = "Ejecutivo Operaciones" Then
                  Puesto = 7
            ElseIf DescPuesto = "Ejecutivo Atencion a Clientes" Then
                  Puesto = 8
            ElseIf DescPuesto = "Director de Ventas" Then
                  Puesto = 9
            ElseIf DescPuesto = "Contabilidad" Then
                  Puesto = 10
            ElseIf DescPuesto = "Ventas y Atencion a Clientes" Then
                  Puesto = 11
             End If
             
            If Val(numero) <> 0 Then
                sqls = "sp_Empleados_Upd  " & numero & ",'" & Nombre & "'," & Puesto & ", 1, '" & DescPuesto & "'"
                cnxbdMty.Execute sqls
            End If
        Next
        Cargar
End With
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub
