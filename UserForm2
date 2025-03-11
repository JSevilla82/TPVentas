

Private Sub ComboBox11_Change()
ListBox7.Clear

If ComboBox11.Text = "Ventas fallidas" Then
Label319 = "ID de interacción"
End If

If ComboBox11.Text = "Ventas exitosas" Then
Label319 = "ID de orden"
End If

End Sub

Private Sub ComboBox3_Change()

On Error GoTo error_Handler:

ComboBox7.Clear

If ComboBox3.Text = "Ventas en espera" Then    '---> Venta en espera

MultiPage6.Value = 0

Label198 = "Fecha de envío"
Label197 = "Agente vendedor"
Label196 = "Nombre del cliente"
Label195 = "Móvil de contacto"
Label194 = "Número documento"
Label193 = "ID de interacción"
Label192 = "Producto a vender"
 
ListBox6.Clear

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT * FROM TB_Principal ORDER BY FECHA_Y_HORA ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ra.EOF
Me.ListBox6.AddItem Ra!FECHA_Y_HORA
Me.ListBox6.List(v, 1) = Ra!USUARIO_AGENTE
Me.ListBox6.List(v, 2) = Ra!NOMBRE_Y_APELLIDO
Me.ListBox6.List(v, 3) = Ra!NUMERO_DOCUMENTO
Me.ListBox6.List(v, 4) = Ra!MOVIL_CONTACTO
Me.ListBox6.List(v, 5) = Ra!ID_ORDEN_INTERACCION
Me.ListBox6.List(v, 6) = Ra!PRODUCTO_VENTA
v = v + 1
Ra.MoveNext
Loop

Label305 = ListBox6.ListCount

Ra.Close
Set Ra = Nothing
miConexion.Close

ComboBox8.Enabled = False
ComboBox8.BackColor = &H80000016

Label299 = "Ventas en espera de ser gestionadas por el Back Office"
Label299.ForeColor = &H0&


End If

'___________________________________________________________________________________________________________________________________________

If ComboBox3.Text = "Ventas en gestión" Then  '---> Venta en gestión

MultiPage6.Value = 1

ListBox6.Clear


Label198 = "Fecha de envío"
Label197 = "Agente vendedor"
Label196 = "Nombre del cliente"
Label195 = "Móvil de contacto"
Label194 = "ID de interacción"
Label193 = "Producto a vender"
Label192 = "Agente gestor"

' Segundo contacto
  
  Call BD_Principal
  Set Ra = New ADODB.Recordset
  Ra.Open "SELECT Count(*) as CONTAR_REGISTROS1 FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  Label313 = Ra("CONTAR_REGISTROS1")
  

  Call BD_Principal
  Set Ra = New ADODB.Recordset
  Ra.Open "SELECT Count(*) as CONTAR_REGISTROS21 FROM TB_Gestionando_Segundo_Contacto", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  Label314 = Ra("CONTAR_REGISTROS21")
  
  Ra.Close
  Set Ra = Nothing
  miConexion.Close
 
 '___________________________________________________________

Label305 = ListBox6.ListCount

Label299 = "Ventas en gestión, elija una opción en el siguiente desplegable"
Label299.ForeColor = &H0&
With ComboBox7
 .AddItem "Nuevas (" & Label313.Caption & ")"
 .AddItem "Agendadas (" & Label314.Caption & ")"
End With
ComboBox8.Enabled = False
ComboBox8.BackColor = &H80000016

End If

'____________________________________________________________________________________________

If ComboBox3.Text = "Ventas agendadas" Then  '----> ventas agendadas

MultiPage6.Value = 0

Label198 = "Fecha de envío"
Label197 = "Agente vendedor"
Label196 = "Nombre del cliente"
Label195 = "Móvil de contacto"
Label194 = "ID de interacción"
Label193 = "Producto a vender"
Label192 = "Fecha a gestionar"

ListBox6.Clear

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT * FROM TB_Segundo_Contacto ORDER BY REANGEDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ra.EOF
Me.ListBox6.AddItem Ra!FECHA_Y_HORA
Me.ListBox6.List(v, 1) = Ra!USUARIO_AGENTE
Me.ListBox6.List(v, 2) = Ra!NOMBRE_Y_APELLIDO
Me.ListBox6.List(v, 3) = Ra!MOVIL_CONTACTO
Me.ListBox6.List(v, 4) = Ra!ID_ORDEN_INTERACCION
Me.ListBox6.List(v, 5) = Ra!PRODUCTO_VENTA
Me.ListBox6.List(v, 6) = Ra!REANGEDADO
v = v + 1
Ra.MoveNext
Loop

Ra.Close
Set Ra = Nothing
miConexion.Close

Label305 = ListBox6.ListCount

Label299 = "Ventas agendadas por el Back Office"
Label299.ForeColor = &H0&

ComboBox8.Enabled = False
ComboBox8.BackColor = &H80000016

End If


Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close
   
End Sub

Private Sub ComboBox7_Change()

On Error GoTo error_Handler:

If ComboBox7.Text = "Nuevas (" & Label313.Caption & ")" Then

ListBox6.Clear

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT * FROM TB_Gestionando ORDER BY FECHA_Y_HORA ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ra.EOF
Me.ListBox6.AddItem Ra!FECHA_Y_HORA
Me.ListBox6.List(v, 1) = Ra!USUARIO_AGENTE
Me.ListBox6.List(v, 2) = Ra!NOMBRE_Y_APELLIDO
Me.ListBox6.List(v, 3) = Ra!MOVIL_CONTACTO
Me.ListBox6.List(v, 4) = Ra!ID_ORDEN_INTERACCION
Me.ListBox6.List(v, 5) = Ra!PRODUCTO_VENTA
Me.ListBox6.List(v, 6) = Ra!AGENTE_GESTOR
v = v + 1
Ra.MoveNext
Loop

Label305 = ListBox6.ListCount

Ra.Close
Set Ra = Nothing
miConexion.Close

Label299 = "Ventas nuevas gestionandose por el Back Office"
Label299.ForeColor = &H0&

End If

'*******************************************************************

If ComboBox7.Text = "Agendadas (" & Label314.Caption & ")" Then

ListBox6.Clear

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT * FROM TB_Gestionando_Segundo_Contacto ORDER BY FECHA_Y_HORA ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ra.EOF
Me.ListBox6.AddItem Ra!FECHA_Y_HORA
Me.ListBox6.List(v, 1) = Ra!USUARIO_AGENTE
Me.ListBox6.List(v, 2) = Ra!NOMBRE_Y_APELLIDO
Me.ListBox6.List(v, 3) = Ra!MOVIL_CONTACTO
Me.ListBox6.List(v, 4) = Ra!ID_ORDEN_INTERACCION
Me.ListBox6.List(v, 5) = Ra!PRODUCTO_VENTA
Me.ListBox6.List(v, 6) = Ra!AGENTE_GESTOR
v = v + 1
Ra.MoveNext
Loop

Label305 = ListBox6.ListCount

Ra.Close
Set Ra = Nothing
miConexion.Close

Label299 = "Ventas agendadas gestionandose por el Back Office"
Label299.ForeColor = &H0&
End If

ComboBox8.Enabled = False
ComboBox8.BackColor = &H80000016

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton25_Click()

On Error GoTo error_Handler:

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText



            Rn.MoveFirst
            Rn.Find " FECHA_REGISTRO = '" & ListBox3.Text & "'"
            
            Ry.AddNew
            Ry!FECHA_REGISTRO = Rn!FECHA_REGISTRO
            Ry!USUARIO_CITRIX = Rn!USUARIO_CITRIX
            Ry!NOMBRE_Y_APELLIDO = Rn!NOMBRE_Y_APELLIDO
            Ry!CONTRASEÑA = Rn!CONTRASEÑA
            Ry!CAMPAÑA = Rn!CAMPAÑA
            Ry.Update
            Rn.Delete
            Rn.MoveNext
            
ListBox3.Clear

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rn.EOF
Me.ListBox3.AddItem Rn!FECHA_REGISTRO
Me.ListBox3.List(d, 1) = Rn!USUARIO_CITRIX
Me.ListBox3.List(d, 2) = Rn!NOMBRE_Y_APELLIDO
Me.ListBox3.List(d, 3) = Rn!CAMPAÑA
d = d + 1
Rn.MoveNext
Loop

ListBox2.Clear

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ry.EOF
Me.ListBox2.AddItem Ry!FECHA_REGISTRO
Me.ListBox2.List(dd, 1) = Ry!USUARIO_CITRIX
Me.ListBox2.List(dd, 2) = Ry!NOMBRE_Y_APELLIDO
Me.ListBox2.List(dd, 3) = Ry!CAMPAÑA
dd = dd + 1
Ry.MoveNext
Loop

Rn.Close
Set Rn = Nothing
Ry.Close
Set Ry = Nothing
miConexion.Close

CommandButton25.Enabled = False
CommandButton25.BackColor = &H80000016
CommandButton28.Enabled = False
CommandButton28.BackColor = &H80000016

Label01 = ListBox2.ListCount
Label02 = ListBox3.ListCount

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close
       
End Sub

Private Sub CommandButton28_Click()

On Error GoTo error_Handler:

Base_Datos = MsgBox("Esta seguro de eliminar este usuario?", vbOKCancel, "TP Ventas vodafone - Panel de control")
If Base_Datos = vbOK Then

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


            Rn.MoveFirst
            Rn.Find " FECHA_REGISTRO = '" & ListBox3.Text & "'"
            Rn.Delete
            Rn.MoveNext
            
            ListBox3.Clear
            
Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rn.EOF
Me.ListBox3.AddItem Rn!FECHA_REGISTRO
Me.ListBox3.List(da, 1) = Rn!USUARIO_CITRIX
Me.ListBox3.List(da, 2) = Rn!NOMBRE_Y_APELLIDO
Me.ListBox3.List(da, 3) = Rn!CAMPAÑA
da = da + 1
Rn.MoveNext
Loop

Rn.Close
Set Rn = Nothing
miConexion.Close

CommandButton25.Enabled = False
CommandButton25.BackColor = &H80000016

CommandButton28.Enabled = False
CommandButton28.BackColor = &H80000016

Label02 = ListBox3.ListCount

End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton29_Click()

On Error GoTo error_Handler:

Base_Datos = MsgBox("Esta seguro de eliminar este usuario?", vbOKCancel, "TP Ventas vodafone - Panel de control")
If Base_Datos = vbOK Then

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Ry.MoveFirst
            Ry.Find " FECHA_REGISTRO = '" & ListBox2.Text & "'"
            Ry.Delete
            Ry.MoveNext


ListBox2.Clear

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ry.EOF
Me.ListBox2.AddItem Ry!FECHA_REGISTRO
Me.ListBox2.List(dd, 1) = Ry!USUARIO_CITRIX
Me.ListBox2.List(dd, 2) = Ry!NOMBRE_Y_APELLIDO
Me.ListBox2.List(dd, 3) = Ry!CAMPAÑA
dd = dd + 1
Ry.MoveNext
Loop


Ry.Close
Set Ry = Nothing
miConexion.Close

CommandButton29.Enabled = False
CommandButton29.BackColor = &H80000016

CommandButton54.Enabled = False
CommandButton54.BackColor = &H80000016

CommandButton30.Enabled = False
CommandButton30.BackColor = &H80000016

Label01 = ListBox2.ListCount


End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close
         
End Sub


Private Sub CommandButton30_Click()


On Error GoTo error_Handler:

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


            Ry.MoveFirst
            Ry.Find " FECHA_REGISTRO = '" & ListBox2.Text & "'"
            
            Rn.AddNew
            Rn!FECHA_REGISTRO = Ry!FECHA_REGISTRO
            Rn!USUARIO_CITRIX = Ry!USUARIO_CITRIX
            Rn!NOMBRE_Y_APELLIDO = Ry!NOMBRE_Y_APELLIDO
            Rn!CONTRASEÑA = Ry!CONTRASEÑA
            Rn!CAMPAÑA = Ry!CAMPAÑA
            Rn.Update
            Ry.Delete
            Ry.MoveNext
            
ListBox3.Clear

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rn.EOF
Me.ListBox3.AddItem Rn!FECHA_REGISTRO
Me.ListBox3.List(d, 1) = Rn!USUARIO_CITRIX
Me.ListBox3.List(d, 2) = Rn!NOMBRE_Y_APELLIDO
Me.ListBox3.List(d, 3) = Rn!CAMPAÑA
d = d + 1
Rn.MoveNext
Loop

ListBox2.Clear

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ry.EOF
Me.ListBox2.AddItem Ry!FECHA_REGISTRO
Me.ListBox2.List(dd, 1) = Ry!USUARIO_CITRIX
Me.ListBox2.List(dd, 2) = Ry!NOMBRE_Y_APELLIDO
Me.ListBox2.List(dd, 3) = Ry!CAMPAÑA
dd = dd + 1
Ry.MoveNext
Loop


Rn.Close
Set Rn = Nothing
Ry.Close
Set Ry = Nothing
miConexion.Close


CommandButton30.Enabled = False
CommandButton30.BackColor = &H80000016
CommandButton29.Enabled = False
CommandButton29.BackColor = &H80000016
CommandButton54.Enabled = False
CommandButton54.BackColor = &H80000016

Label01 = ListBox2.ListCount
Label02 = ListBox3.ListCount

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton39_Click()

MultiPage1.Value = 0
txtUsuario = ""
txtPassword = ""
txtUsuario.SetFocus
End Sub

Private Sub CommandButton46_Click()

On Error GoTo error_Handler:

If txtUsuario.Text = "" Or txtPassword.Text = "" Then
MsgBox "Para validar su identidad debe ingresar un nombre de usuario y una contraseña. ", vbInformation, "Inicio de sesión"
Exit Sub

Else

Label435.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "select * from Usuarios_Admin where USUARIO_CITRIX ='" + txtUsuario.Text + "' and CONTRASEÑA='" + txtPassword.Text + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

If Rn.EOF = True Then
Label435.ForeColor = &HFFFFFF
MsgBox "El nombre de usuario y la contraseña que ingresaste no coinciden con nuestros registros. Por favor, revisa e inténtalo de nuevo.", vbCritical, "Inicio de sesión"

txtUsuario = ""
txtPassword = ""
txtUsuario.SetFocus

Rn.Close
Set Rn = Nothing
miConexion.Close
Else

Rn.Close
Set Rn = Nothing
miConexion.Close

Label435.ForeColor = &HFFFFFF

UserForm1.Caption = "TP Ventas vodafone - Panel de control"
MultiPage1.Value = 1

End If
End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton49_Click()

On Error GoTo error_Handler:

Base_Datos = MsgBox("Esta seguro de eliminar este usuario?", vbOKCancel, "TP Ventas vodafone - Panel de control")
If Base_Datos = vbOK Then


Call BD_Seguridad
Set Rp = New ADODB.Recordset
Rp.Open "SELECT * FROM Usuarios_agentes", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rp.MoveFirst
Rp.Find " FECHA_REGISTRO = '" & ListBox5.Text & "'"
Rp.Delete
Rp.MoveNext


ListBox5.Clear

Call BD_Seguridad
Set Rp = New ADODB.Recordset
Rp.Open "SELECT * FROM Usuarios_agentes", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rp.EOF
Me.ListBox5.AddItem Rp!FECHA_REGISTRO
Me.ListBox5.List(p, 1) = Rp!USUARIO_CITRIX
Me.ListBox5.List(p, 2) = Rp!NOMBRE_Y_APELLIDO
Me.ListBox5.List(p, 3) = Rp!CAMPAÑA
p = p + 1
Rp.MoveNext
Loop


Rp.Close
Set Rp = Nothing
miConexion.Close

CommandButton49.Enabled = False
CommandButton49.BackColor = &H80000016
CommandButton62.Enabled = False
CommandButton62.BackColor = &H80000016

Label03 = ListBox5.ListCount

End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton54_Click()
MultiPage1.Value = 4
MultiPage5.Value = 1
End Sub

Private Sub CommandButton55_Click()


On Error GoTo error_Handler:

ListBox2.Clear

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ry.EOF
Me.ListBox2.AddItem Ry!FECHA_REGISTRO
Me.ListBox2.List(dd, 1) = Ry!USUARIO_CITRIX
Me.ListBox2.List(dd, 2) = Ry!NOMBRE_Y_APELLIDO
Me.ListBox2.List(dd, 3) = Ry!CAMPAÑA
dd = dd + 1
Ry.MoveNext
Loop

Ry.Close
Set Ry = Nothing
miConexion.Close

CommandButton29.Enabled = False
CommandButton29.BackColor = &H80000016

CommandButton54.Enabled = False
CommandButton54.BackColor = &H80000016

CommandButton30.Enabled = False
CommandButton30.BackColor = &H80000016

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton56_Click()


On Error GoTo error_Handler:

ListBox3.Clear

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rn.EOF
Me.ListBox3.AddItem Rn!FECHA_REGISTRO
Me.ListBox3.List(d, 1) = Rn!USUARIO_CITRIX
Me.ListBox3.List(d, 2) = Rn!NOMBRE_Y_APELLIDO
Me.ListBox3.List(d, 3) = Rn!CAMPAÑA
d = d + 1
Rn.MoveNext
Loop

Rn.Close
Set Rn = Nothing
miConexion.Close

CommandButton28.Enabled = False
CommandButton28.BackColor = &H80000016

CommandButton25.Enabled = False
CommandButton25.BackColor = &H80000016

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton62_Click()

MultiPage1.Value = 4
MultiPage5.Value = 0

End Sub

Private Sub CommandButton63_Click()


On Error GoTo error_Handler:

ListBox5.Clear

Call BD_Seguridad
Set Rp = New ADODB.Recordset
Rp.Open "SELECT * FROM Usuarios_agentes ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rp.EOF
Me.ListBox5.AddItem Rp!FECHA_REGISTRO
Me.ListBox5.List(p, 1) = Rp!USUARIO_CITRIX
Me.ListBox5.List(p, 2) = Rp!NOMBRE_Y_APELLIDO
Me.ListBox5.List(p, 3) = Rp!CAMPAÑA
p = p + 1
Rp.MoveNext
Loop


Rp.Close
Set Rp = Nothing
miConexion.Close

CommandButton49.Enabled = False
CommandButton49.BackColor = &H80000016
CommandButton62.Enabled = False
CommandButton62.BackColor = &H80000016

Label03 = ListBox5.ListCount

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub





Private Sub CommandButton64_Click()


On Error GoTo error_Handler:

If TextBox1.Text <> TextBox2.Text Then
MsgBox "Las contraseñas no coinciden. Vuelve a intentarlo ", vbCritical, "TP Ventas vodafone - Panel de control"
TextBox1 = ""
TextBox2 = ""
TextBox1.SetFocus
Exit Sub

Else

Base_Datos = MsgBox("Esta seguro de terminar y enviar esta información?", vbOKCancel, "TP Ventas vodafone - Panel de control")
If Base_Datos = vbOK Then

Call BD_Seguridad
Set Rp = New ADODB.Recordset
Rp.Open "SELECT * FROM Usuarios_agentes", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rp.MoveFirst
Rp.Find " FECHA_REGISTRO = '" & ListBox5.Text & "'"
Rp!CONTRASEÑA = TextBox1.Text
Rp.Update


If Rp.State = 1 Or Rp.State = 0 Then
MsgBox "El cambio de contraseña se ha realizado correctamente.", vbInformation, "TP Ventas vodafone - Panel de control"

TextBox1.Text = ""
TextBox2.Text = ""
MultiPage1.Value = 3

Rp.Close
Set Rp = Nothing
miConexion.Close

Else
MsgBox "Ha ocurrido un error inesperado, no se ha podido realizar el cambio de contraseña.", vbCritical, "TP Ventas vodafone - Panel de control"

Rp.Close
Set Rp = Nothing
miConexion.Close

End If
End If
End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton65_Click()

On Error GoTo error_Handler:

If TextBox3.Text <> TextBox4.Text Then
MsgBox "Las contraseñas no coinciden. Vuelve a intentarlo ", vbCritical, "TP Ventas vodafone - Panel de control"
TextBox3 = ""
TextBox4 = ""
TextBox3.SetFocus
Exit Sub

Else

Base_Datos = MsgBox("Esta seguro de terminar y enviar esta información?", vbOKCancel, "TP Ventas vodafone - Panel de control")
If Base_Datos = vbOK Then

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Ry.MoveFirst
Ry.Find " FECHA_REGISTRO = '" & ListBox2.Text & "'"
Ry!CONTRASEÑA = TextBox3.Text
Ry.Update

If Ry.State = 1 Or Ry.State = 0 Then
MsgBox "El cambio de contraseña se ha realizado correctamente.", vbInformation, "TP Ventas vodafone - Panel de control"

TextBox3.Text = ""
TextBox4.Text = ""
MultiPage1.Value = 2

Ry.Close
Set Ry = Nothing
miConexion.Close

Else
MsgBox "Ha ocurrido un error inesperado, no se ha podido realizar el cambio de contraseña.", vbCritical, "TP Ventas vodafone - Panel de control"

Ry.Close
Set Ry = Nothing
miConexion.Close

End If
End If
End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton66_Click()
MultiPage1.Value = 1
End Sub

Private Sub CommandButton67_Click()

On Error GoTo error_Handler:

CommandButton67.Enabled = False
CommandButton67.BackColor = &H80000016
CommandButton67.Caption = "Buscando usuario..."
Application.Wait (Now + TimeValue("00:00:01"))

ListBox5.Clear

Call BD_Seguridad
Set Rp = New ADODB.Recordset
Rp.Open "SELECT * FROM Usuarios_agentes WHERE USUARIO_CITRIX LIKE '%" & TextBox5.Text & "%'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rp.EOF = False
Me.ListBox5.AddItem Rp!FECHA_REGISTRO
Me.ListBox5.List(p, 1) = Rp!USUARIO_CITRIX
Me.ListBox5.List(p, 2) = Rp!NOMBRE_Y_APELLIDO
Me.ListBox5.List(p, 3) = Rp!CAMPAÑA
p = p + 1
Rp.MoveNext
Wend

Label03 = ListBox5.ListCount

Rp.Close
Set Rp = Nothing
miConexion.Close

CommandButton67.Enabled = True
CommandButton67.BackColor = &HC0&
CommandButton67.Caption = "Buscar usuario"

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton69_Click()
MultiPage1.Value = 1
End Sub

Private Sub CommandButton70_Click()

On Error GoTo error_Handler:

ListBox7.Clear

If Not IsDate(TextBox80.Text) Then
MsgBox "Error en la consulta, el formato de fecha no es válido", vbInformation, "TP Ventas vodafone - Panel de control"
Exit Sub
Else
If Not IsDate(TextBox81.Text) Then
MsgBox "Error en la consulta, el formato de fecha no es válido", vbInformation, "TP Ventas vodafone - Panel de control"
Exit Sub
Else

If ComboBox11.Text = "Ventas exitosas" Then
If ComboBox1.Text = "Todos los usuarios" Then
    
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count (*) AS CONTAR_REGISTROS1 FROM TB_Gestionados_Exitosas WHERE FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label409 = Rg("CONTAR_REGISTROS1")

If Label409.Caption = "0" Then
Label316 = "La consulta en la base de datos no arrojó ningún resultado"
Label316.ForeColor = &H0&
ListBox7.Clear
Rg.Close
Set Rg = Nothing
miConexion.Close
Exit Sub

Else

ListBox7.Clear
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT * FROM TB_Gestionados_Exitosas WHERE FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rg.EOF
Me.ListBox7.AddItem Rg!FECHA_Y_HORA
Me.ListBox7.List(v, 1) = Rg!AGENTE_GESTOR
Me.ListBox7.List(v, 2) = Rg!NOMBRE_Y_APELLIDO
Me.ListBox7.List(v, 3) = Rg!NUMERO_DOCUMENTO
Me.ListBox7.List(v, 4) = Rg!FECHA_FIN_GESTION
Me.ListBox7.List(v, 5) = Rg!ID_ORDEN
Me.ListBox7.List(v, 6) = Rg!RESULTADO_FINAL
v = v + 1
Rg.MoveNext
Loop

Rg.Close
Set Rg = Nothing
miConexion.Close

Label316 = "Ventas gestionadas por el Back Office con resultado exitoso"
Label316.ForeColor = &H0&
End If
    
Else
    
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count (*) AS CONTAR_REGISTROS1 FROM TB_Gestionados_Exitosas WHERE AGENTE_GESTOR = '" & ComboBox1.Text & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label409 = Rg("CONTAR_REGISTROS1")

If Label409.Caption = "0" Then
Label316 = "La consulta en la base de datos no arrojó ningún resultado"
Label316.ForeColor = &H0&
ListBox7.Clear

Rg.Close
Set Rg = Nothing
miConexion.Close

Exit Sub
    
Else
   
ListBox7.Clear
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT * FROM TB_Gestionados_Exitosas WHERE AGENTE_GESTOR = '" & ComboBox1.Text & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rg.EOF
Me.ListBox7.AddItem Rg!FECHA_Y_HORA
Me.ListBox7.List(v, 1) = Rg!AGENTE_GESTOR
Me.ListBox7.List(v, 2) = Rg!NOMBRE_Y_APELLIDO
Me.ListBox7.List(v, 3) = Rg!NUMERO_DOCUMENTO
Me.ListBox7.List(v, 4) = Rg!FECHA_FIN_GESTION
Me.ListBox7.List(v, 5) = Rg!ID_ORDEN
Me.ListBox7.List(v, 6) = Rg!RESULTADO_FINAL
v = v + 1
Rg.MoveNext
Loop

Rg.Close
Set Rg = Nothing
miConexion.Close
    
Label316 = "Registros gestionados por " & ComboBox1.Text & " con resultado exitoso"
Label316.ForeColor = &H0&
    
End If
End If
End If
End If
End If
    
'///////////////////////////////////////////////////////////////////


If ComboBox11.Text = "Ventas fallidas" Then
If ComboBox1.Text = "Todos los usuarios" Then
    
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count (*) AS CONTAR_REGISTROS1 FROM TB_Gestionados_Fallidas WHERE FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label409 = Rg("CONTAR_REGISTROS1")

If Label409.Caption = "0" Then
Label316 = "La consulta en la base de datos no arrojó ningún resultado"
Label316.ForeColor = &H0&
ListBox7.Clear
Rg.Close
Set Rg = Nothing
miConexion.Close
Exit Sub

Else

ListBox7.Clear
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT * FROM TB_Gestionados_Fallidas WHERE FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rg.EOF
Me.ListBox7.AddItem Rg!FECHA_Y_HORA
Me.ListBox7.List(v, 1) = Rg!AGENTE_GESTOR
Me.ListBox7.List(v, 2) = Rg!NOMBRE_Y_APELLIDO
Me.ListBox7.List(v, 3) = Rg!NUMERO_DOCUMENTO
Me.ListBox7.List(v, 4) = Rg!FECHA_FIN_GESTION
Me.ListBox7.List(v, 5) = Rg!ID_ORDEN_INTERACCION
Me.ListBox7.List(v, 6) = Rg!RESULTADO_FINAL
v = v + 1
Rg.MoveNext
Loop

Rg.Close
Set Rg = Nothing
miConexion.Close

Label316 = "Ventas gestionadas por el Back Office con resultado fallido"
Label316.ForeColor = &H0&
End If
    
Else
    
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count (*) AS CONTAR_REGISTROS1 FROM TB_Gestionados_Fallidas WHERE AGENTE_GESTOR = '" & ComboBox1.Text & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label409 = Rg("CONTAR_REGISTROS1")

If Label409.Caption = "0" Then
Label316 = "La consulta en la base de datos no arrojó ningún resultado"
Label316.ForeColor = &H0&
ListBox7.Clear

Rg.Close
Set Rg = Nothing
miConexion.Close

Exit Sub
    
Else
   
ListBox7.Clear
Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT * FROM TB_Gestionados_Fallidas WHERE AGENTE_GESTOR = '" & ComboBox1.Text & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox80.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rg.EOF
Me.ListBox7.AddItem Rg!FECHA_Y_HORA
Me.ListBox7.List(v, 1) = Rg!AGENTE_GESTOR
Me.ListBox7.List(v, 2) = Rg!NOMBRE_Y_APELLIDO
Me.ListBox7.List(v, 3) = Rg!NUMERO_DOCUMENTO
Me.ListBox7.List(v, 4) = Rg!FECHA_FIN_GESTION
Me.ListBox7.List(v, 5) = Rg!ID_ORDEN_INTERACCION
Me.ListBox7.List(v, 6) = Rg!RESULTADO_FINAL
v = v + 1
Rg.MoveNext
Loop

Rg.Close
Set Rg = Nothing
miConexion.Close
    
Label316 = "Registros gestionados por " & ComboBox1.Text & " con resultado fallido"
Label316.ForeColor = &H0&
    
End If
End If
End If


Label326 = ListBox7.ListCount
  
  
Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton76_Click()

On Error GoTo error_Handler:

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count (USUARIO_AGENTE) AS CONTAR_REGISTROS1 FROM TB_Gestionados_Exitosas WHERE FECHA_FIN_GESTION BETWEEN #" & TextBox82.Text & "# AND #" & Label411.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label383 = Rg("CONTAR_REGISTROS1")

Rg.Close
Set Rg = Nothing
miConexion.Close

If Label383.Caption = 0 Then
MsgBox "No se encontraron datos para la exportación", vbInformation, "Exportar ventas - Exitosas"
Exit Sub

Else
  
Base_Datos = MsgBox("Esta seguro de comenzar con la expotación de " & Label383.Caption & " ventas?", vbOKCancel, "Exportar ventas - Exitosas")
If Base_Datos = vbOK Then

Label437.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))
        
'________________________________________________________________________

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT * FROM TB_Gestionados_Exitosas WHERE FECHA_FIN_GESTION BETWEEN #" & TextBox82.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_FIN_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

NombreHoja = "Ventas exitosas"

Set APIExcel = CreateObject("Excel.Application")
Set AddLibro = APIExcel.Workbooks.Add
APIExcel.Visible = False

Set AddHoja = AddLibro.Worksheets(1)
If Len(NombreHoja) > 0 Then AddHoja.Name = Left(NombreHoja, 30)
columnas = Rg.Fields.Count
For i = 0 To columnas - 1
APIExcel.Cells(1, i + 1) = Rg.Fields(i).Name
Next i

Rg.MoveFirst
AddHoja.Range("A2").CopyFromRecordset Rg

With APIExcel.ActiveSheet.Cells
.Select
.EntireColumn("L").Delete
'.EntireColumn("O").ClearFormats
.EntireColumn("A:Q").AutoFit
.Range("A1").Select
'.Cells.ClearFormats
End With

'APIExcel.Application.ActiveWorkbook.SaveAs Filename:="J:\OTROS\Formaci\GEOS VF - DATOS EXPORTADOS\Geos vodafone - Exitosas\R.E - Puerta del sol  " & Format(Now, "DD-MMM-YYYY hh-mm-ss") & ".xlsx"

Rg.Close
Set Rg = Nothing
miConexion.Close

Label437.ForeColor = &HFFFFFF

Base_Datos = MsgBox("La expotación de " & Label383.Caption & " ventas ha finalizado con éxito, desea abrir el contenido exportado?", vbOKCancel, "Exportar ventas - Exitosas")
If Base_Datos = vbOK Then
       
APIExcel.Visible = True

End If
End If
End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Ry = Nothing
miConexion.Close

End Sub

Private Sub CommandButton77_Click()

'On Error GoTo error_Handler:

Base_Datos = MsgBox("Esta seguro de comenzar con la expotación de todas las ventas?", vbOKCancel, "Exportar ventas")
If Base_Datos = vbOK Then

Label436.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

Call BD_Principal
Set Rg = New ADODB.Recordset
Rg.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,TIPO_ALTA,PRODUCTO_CARACTERISTICA,OFERTA_ASOCIADA FROM TB_Principal WHERE FECHA_ENVIO BETWEEN #" & TextBox84.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,TIPO_ALTA,PRODUCTO_CARACTERISTICA,OFERTA_ASOCIADA FROM TB_Gestionando WHERE FECHA_ENVIO BETWEEN #" & TextBox84.Text & "# AND #" & Label411.Caption & "# ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


NombreHoja_1 = "Bandeja de entrada"
NombreHoja_2 = "Ventas en gestión"

Set APIExcel = CreateObject("Excel.Application")
Set AddLibro = APIExcel.Workbooks.Add
APIExcel.Visible = False


Set AddHoja = AddLibro.Worksheets(1)
If Len(NombreHoja_1) > 0 Then AddHoja.Name = Left(NombreHoja_1, 30)
columnas = Rg.Fields.Count
For i = 0 To columnas - 1
APIExcel.Cells(1, i + 1) = Rg.Fields(i).Name
Next i
Rg.MoveFirst
AddHoja.Range("A2").CopyFromRecordset Rg

columnass = Rs.Fields.Count
For F = 0 To columnass - 1
APIExcel.Cells(1, F + 1) = Rs.Fields(F).Name
Next F
Rs.MoveFirst
AddHoja.Range("A2").CopyFromRecordset Rs


'With APIExcel.ActiveSheet.Cells
'.Select
'.EntireColumn("A:I").AutoFit
'.Range("A1").Select
'End With

Rg.Close
Set Rg = Nothing
Rs.Close
Set Rs = Nothing
miConexion.Close
    
Base_Datos = MsgBox("La expotación de ventas ha finalizado con éxito, desea abrir el contenido exportado?", vbOKCancel, "Exportar ventas")
If Base_Datos = vbOK Then
APIExcel.Visible = True
End If
    
End If

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Ry = Nothing
miConexion.Close

End Sub





Private Sub Image30_Click()
MultiPage1.Value = 1
End Sub

Private Sub Image31_Click()
MultiPage1.Value = 1
End Sub

Private Sub Image36_Click()

Label350 = "Fecha inicio"
Load frmCalendario
frmCalendario.Show

End Sub

Private Sub Image37_Click()
Label350 = "Fecha fin"
Load frmCalendario
frmCalendario.Show
End Sub

Private Sub Image50_Click()
MultiPage1.Value = 1
End Sub

Private Sub Image59_Click()
Label350 = "Fecha inicio FL"
Load frmCalendario
frmCalendario.Show
End Sub

Private Sub Image60_Click()
Label350 = "Fecha fin FL"
Load frmCalendario
frmCalendario.Show
End Sub

Private Sub Image61_Click()
Label350 = "Fecha inicio EX"
Load frmCalendario
frmCalendario.Show
End Sub

Private Sub Image62_Click()
Label350 = "Fecha fin EX"
Load frmCalendario
frmCalendario.Show
End Sub

Private Sub Label221_Click()
MultiPage1.Value = 3
End Sub

Private Sub Label247_Click()

On Error GoTo error_Handler:

Label438.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

CommandButton29.Enabled = False
CommandButton29.BackColor = &H80000016
CommandButton30.Enabled = False
CommandButton30.BackColor = &H80000016
CommandButton25.Enabled = False
CommandButton25.BackColor = &H80000016
CommandButton28.Enabled = False
CommandButton28.BackColor = &H80000016

CommandButton54.Enabled = False
CommandButton54.BackColor = &H80000016



'__________________________________________________________________________

ListBox3.Clear

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rn.EOF
Me.ListBox3.AddItem Rn!FECHA_REGISTRO
Me.ListBox3.List(d, 1) = Rn!USUARIO_CITRIX
Me.ListBox3.List(d, 2) = Rn!NOMBRE_Y_APELLIDO
Me.ListBox3.List(d, 3) = Rn!CAMPAÑA
d = d + 1
Rn.MoveNext
Loop

Rn.Close
Set Rn = Nothing


ListBox2.Clear


Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ry.EOF
Me.ListBox2.AddItem Ry!FECHA_REGISTRO
Me.ListBox2.List(dd, 1) = Ry!USUARIO_CITRIX
Me.ListBox2.List(dd, 2) = Ry!NOMBRE_Y_APELLIDO
Me.ListBox2.List(dd, 3) = Ry!CAMPAÑA
dd = dd + 1
Ry.MoveNext
Loop


Ry.Close
Set Ry = Nothing
miConexion.Close

Label01 = ListBox2.ListCount
Label02 = ListBox3.ListCount

Label438.ForeColor = &HFFFFFF

MultiPage1.Value = 2

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Ry = Nothing
miConexion.Close

End Sub

Private Sub Label248_Click()

On Error GoTo error_Handler:

Label438.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

ListBox5.Clear

Call BD_Seguridad
Set Rp = New ADODB.Recordset
Rp.Open "SELECT * FROM Usuarios_agentes ORDER BY FECHA_REGISTRO DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rp.EOF
Me.ListBox5.AddItem Rp!FECHA_REGISTRO
Me.ListBox5.List(p, 1) = Rp!USUARIO_CITRIX
Me.ListBox5.List(p, 2) = Rp!NOMBRE_Y_APELLIDO
Me.ListBox5.List(p, 3) = Rp!CAMPAÑA
p = p + 1
Rp.MoveNext
Loop

Rp.Close
Set Rp = Nothing
miConexion.Close

CommandButton49.Enabled = False
CommandButton49.BackColor = &H80000016
CommandButton62.Enabled = False
CommandButton62.BackColor = &H80000016
CommandButton67.Enabled = False
CommandButton67.BackColor = &H80000016

Label03 = ListBox5.ListCount

Label438.ForeColor = &HFFFFFF
MultiPage1.Value = 3

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Label251_Click()

'On Error GoTo error_Handler:

Exit Sub

Label438.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

ListBox6.Clear

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT * FROM TB_Principal ORDER BY FECHA_Y_HORA ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ra.EOF
Me.ListBox6.AddItem Ra!FECHA_Y_HORA
Me.ListBox6.List(v, 1) = Ra!USUARIO_AGENTE
Me.ListBox6.List(v, 2) = Ra!NOMBRE_Y_APELLIDO
Me.ListBox6.List(v, 3) = Ra!NUMERO_DOCUMENTO
Me.ListBox6.List(v, 4) = Ra!MOVIL_CONTACTO
Me.ListBox6.List(v, 5) = Ra!ID_ORDEN_INTERACCION
Me.ListBox6.List(v, 6) = Ra!PRODUCTO_VENTA
v = v + 1
Ra.MoveNext
Loop


Label305 = ListBox6.ListCount
ComboBox3.Text = "Ventas en espera"

CommandButton42.Enabled = False
CommandButton42.BackColor = &H80000016

'_____________________________________ Contar datos estado de gestion!________________________________________
  
  Call BD_Principal
  Set Ra = New ADODB.Recordset
  Ra.Open "SELECT Count(*) as CONTAR_REGISTROS FROM TB_Principal", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
 Label283 = Ra("CONTAR_REGISTROS")

    Call BD_Principal
  Set Ra = New ADODB.Recordset
  Ra.Open "SELECT Count(*) as CONTAR_REGISTROS2 FROM TB_Segundo_Contacto", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  Label306 = Ra("CONTAR_REGISTROS2")
  
  
  ' Segundo contacto
  
  Call BD_Principal
  Set Ra = New ADODB.Recordset
  Ra.Open "SELECT Count(*) as CONTAR_REGISTROS1 FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  Label313 = Ra("CONTAR_REGISTROS1")
  

  Call BD_Principal
  Set Ra = New ADODB.Recordset
  Ra.Open "SELECT Count(*) as CONTAR_REGISTROS21 FROM TB_Gestionando_Segundo_Contacto", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
  Label314 = Ra("CONTAR_REGISTROS21")
  
  Label308.Caption = Val(Label313.Caption) + Val(Label314.Caption)
  
  Label311.Caption = Val(Label283.Caption) + Val(Label306.Caption) + Val(Label308.Caption)
  

   Label299 = "Ventas en espera de ser gestionadas por el Back Office"
   Label299.ForeColor = &H0&
   
   

Ra.Close
Set Ra = Nothing
miConexion.Close

Label438.ForeColor = &HFFFFFF

MultiPage1.Value = 6
MultiPage6.Value = 0

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Label255_Click()

Exit Sub

Label438.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

Label438.ForeColor = &HFFFFFF

MultiPage1.Value = 8
MultiPage8.Value = 0

TextBox82 = ""
TextBox83 = ""

End Sub

Private Sub Label295_Click()
 MultiPage1.Value = 2
End Sub

Private Sub Label302_Click()

Exit Sub
On Error GoTo error_Handler:

Label438.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

ComboBox11.Text = "Ventas Exitosas"

Call BD_Seguridad
Set Ru = New ADODB.Recordset
Ru.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Ru.EOF
ComboBox1.AddItem Ru.Fields(2)
Ru.MoveNext
Loop

Ru.Close
Set Ru = Nothing

Call BD_Gestionados
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(*) as CONTAR_REGISTROS1 FROM TB_Gestionados_Exitosas", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label331 = Ra("CONTAR_REGISTROS1")

Call BD_Gestionados
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(*) as CONTAR_REGISTROS21 FROM TB_Gestionados_Fallidas", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label329 = Ra("CONTAR_REGISTROS21")

Ra.Close
Set Ra = Nothing
miConexion.Close
  
Label327.Caption = Val(Label331.Caption) + Val(Label329.Caption)

ComboBox12.Text = "Filtrado manual"

ComboBox12.Enabled = False
ComboBox12.BackColor = &H80000016
  
Label316 = "Nota: Para consultar un dia en especifico debe colocar la misma fecha en fecha inicio y fecha fin"
Label316.ForeColor = &H0&

Label438.ForeColor = &HFFFFFF
MultiPage1.Value = 7
MultiPage7.Value = 0

Exit Sub
error_Handler:
MultiPage1.Value = 5
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = txtUsuario.Text
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close



End Sub

Private Sub Label446_Click()

Exit Sub
Label438.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

Label438.ForeColor = &HFFFFFF

MultiPage1.Value = 8
MultiPage8.Value = 1

TextBox84 = ""
TextBox85 = ""

End Sub

Private Sub ListBox2_Click()

CommandButton30.Enabled = True
CommandButton30.BackColor = &HC0&

CommandButton54.Enabled = True
CommandButton54.BackColor = &HC0&

CommandButton29.Enabled = True
CommandButton29.BackColor = &HC0&

CommandButton25.Enabled = False
CommandButton25.BackColor = &H80000016

End Sub

Private Sub ListBox3_Click()
CommandButton25.Enabled = True
CommandButton25.BackColor = &HC0&

CommandButton28.Enabled = True
CommandButton28.BackColor = &HC0&

CommandButton30.Enabled = False
CommandButton30.BackColor = &H80000016
End Sub


Private Sub ListBox5_Click()
CommandButton49.Enabled = True
CommandButton49.BackColor = &HC0&

CommandButton62.Enabled = True
CommandButton62.BackColor = &HC0&
End Sub


Private Sub ListBox6_Click()
ComboBox8.Enabled = True
ComboBox8.BackColor = &HC0&
End Sub


Private Sub TextBox5_Change()
TextBox5 = UCase(TextBox5)
CommandButton67.Enabled = True
CommandButton67.BackColor = &HC0&
End Sub

Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
        CommandButton67.SetFocus
        CommandButton67_Click
        End If
End Sub

Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
        CommandButton46.SetFocus
        CommandButton46_Click
        End If
End Sub

Private Sub txtUsuario_Change()
txtUsuario = UCase(txtUsuario)
End Sub

Private Sub txtUsuario_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
        CommandButton46.SetFocus
        CommandButton46_Click
        End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Base_Datos = MsgBox("¿Seguro que quieres salir?", vbOKCancel, "TP Ventas vodafone")

If Base_Datos = vbOK Then
Unload Me
UserForm2.Hide
End If
Cancel = True
End Sub

Private Sub UserForm_Initialize()

MultiPage1.Value = 0

  With ComboBox3
 .AddItem "Ventas en espera"
 .AddItem "Ventas en gestión"
 .AddItem "Ventas agendadas"
 End With
 
 With ComboBox11
 .AddItem "Ventas exitosas"
 .AddItem "Ventas fallidas"
 End With
 
 With ComboBox1
 .AddItem "Todos los usuarios"
 End With
 
  With ComboBox12
 .AddItem "Filtrado manual"
 End With
 
CommandButton52.Enabled = False
CommandButton52.BackColor = &H80000016

UserForm2.Caption = "TP Ventas vodafone - Panel de control"
End Sub
