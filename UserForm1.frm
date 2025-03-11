Dim Gestión_Time(1 To 3) As Integer

Dim Inicio_Gestión As Date
Dim Fin_Gestión As Date
Dim Total_Gestión As Date

Sub Error_sistemas() 'Este codigo es ejecutado despues de un error.

On Error GoTo error_Handler2:

Call BD_Seguridad
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rs.AddNew
Rs.Fields("USUARIO") = Environ("USERNAME")
Rs.Fields("DESCRIPCION") = Error_Sistema
Rs.Fields("APLICACION") = "TP Ventas vodafone - Gestión Back Office"
Rs.Update

Rs.Close
miConexion.Close

MultiPage1.Value = 4
Application.Wait (Now + TimeValue("00:00:01"))
'CommandButton39_Click

Exit Sub
error_Handler2:

MultiPage1.Value = 4
Application.Wait (Now + TimeValue("00:00:01"))
'CommandButton39_Click

End Sub


Private Sub ComboBox004_Change()

On Error GoTo error_Handler:

ComboBox005.Clear

Call BD_Tipificacion
Set Rs = New ADODB.Recordset
Rs.Open "SELECT DISTINCT(TIPO_ALTA) FROM Productos_items WHERE TIPO_VENTA = '" & ComboBox004.Text & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
ComboBox005.AddItem Rs.Fields(0)
Rs.MoveNext
Loop

Rs.Close
miConexion.Close

Label462.Visible = False

Exit Sub
error_Handler:
Error_Sistema = Err.Description
MultiPage1.Value = 3
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"
Application.Wait (Now + TimeValue("00:00:03"))

Error_sistemas

End Sub

Private Sub ComboBox005_Change()

On Error GoTo error_Handler:

ComboBox006.Clear

Call BD_Tipificacion
Set Rs = New ADODB.Recordset
Rs.Open "SELECT DISTINCT(PRODUCTO) FROM Productos_items WHERE TIPO_VENTA = '" & ComboBox004.Text & "'" & "and TIPO_ALTA = '" & ComboBox005.Text & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
ComboBox006.AddItem Rs.Fields(0)
Rs.MoveNext
Loop

Rs.Close
miConexion.Close

Label462.Visible = False

Exit Sub
error_Handler:
Error_Sistema = Err.Description
MultiPage1.Value = 3
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"
Application.Wait (Now + TimeValue("00:00:03"))

Error_sistemas

End Sub

Private Sub ComboBox006_Change()

On Error GoTo error_Handler:

ComboBox007.Clear

Call BD_Tipificacion
Set Rs = New ADODB.Recordset
Rs.Open "SELECT DISTINCT(OFERTA_ASOCIADA) FROM Productos_items WHERE TIPO_VENTA = '" & ComboBox004.Text & "'" & _
            " and TIPO_ALTA = '" & ComboBox005.Text & "'" & _
            " and PRODUCTO = '" & ComboBox006.Text & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
ComboBox007.AddItem Rs.Fields(0)
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing

miConexion.Close

Label462.Visible = False

Exit Sub
error_Handler:
Error_Sistema = Err.Description
MultiPage1.Value = 3
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"
Application.Wait (Now + TimeValue("00:00:03"))

Error_sistemas

End Sub

Private Sub ComboBox007_Change()
Label462.Visible = False
End Sub

Private Sub ComboBox1_Change()

ComboBox2.Enabled = True
ComboBox2.BackColor = &H80000018
ComboBox2.Clear

If ComboBox1 = "Venta en curso" Then
With ComboBox2
.AddItem "Pendiente completar"
End With
End If

If ComboBox1 = "Agendar venta" Then
With ComboBox2
.AddItem "Cliente agenda"
.AddItem "Corrección de datos"
.AddItem "Cliente ilocalizable"
.AddItem "OT pendiente cancelar"
End With
End If

If ComboBox1 = "Venta exitosa" Then
With ComboBox2
.AddItem "Terminal recibido"
.AddItem "Servicio instalado"
.AddItem "Tarifa provisionada"
.AddItem "Producto provisionado"
End With
End If

If ComboBox1 = "Venta fallida" Then
With ComboBox2
.AddItem "Cliente cancela"
.AddItem "Oferta incorrecta"
.AddItem "Imposible instalar"
.AddItem "Contraoferta externa"
.AddItem "Cliente ilocalizable"

End With
End If

TextBox95 = ""
TextBox95.Enabled = False
TextBox95.BackColor = &HE0E0E0

TextBox6 = ""
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0

TextBox2 = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0

TextBox5 = ""
TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0

TextBox2 = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0

Label477.Visible = False
TextBox95.Visible = False

MultiPage3.Value = 0
ComboBox2.SetFocus

End Sub

Private Sub ComboBox2_Change()

If ComboBox1 = "Venta en curso" Then
MultiPage3.Value = 0
TextBox5 = ""
TextBox5.Enabled = True
TextBox5.BackColor = &H80000018
TextBox5.SetFocus
TextBox2.Text = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
Label477.Visible = True
TextBox95.Visible = True
End If

If ComboBox1 = "Agendar venta" Then
MultiPage3.Value = 1
ComboBox3 = ""
ComboBox3.Enabled = True
ComboBox3.BackColor = &H80000018
ComboBox3.SetFocus
TextBox6 = ""
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0
TextBox2.Text = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
End If

If ComboBox1 = "Venta exitosa" Or ComboBox1 = "Venta fallida" Then
If ComboBox1 = "Venta fallida" Then
MultiPage3.Value = 0
TextBox5 = ""
TextBox5 = "No requerido"
TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0
TextBox2.SetFocus
Else
MultiPage3.Value = 0
TextBox5 = ""
TextBox5.Enabled = True
TextBox5.BackColor = &H80000018
TextBox5.SetFocus

TextBox2.Text = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
End If
End If

End Sub

Private Sub ComboBox3_Change()
TextBox6.Enabled = True
TextBox6.BackColor = &H80000018
TextBox6 = ""
TextBox6 = Now
End Sub


Private Sub ComboBox8_Change()
Label474.Visible = False
'TextBox94.SetFocus
End Sub

Private Sub CommandButton11_Click()

'On Error GoTo error_Handler

ListBox1.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Principal ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox1.AddItem Rs!ID_REGISTRO
Me.ListBox1.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox1.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox1.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox1.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox1.List(a, 5) = Rs!FECHA_ENVIO
a = a + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing
miConexion.Close

Label14 = ListBox1.ListCount
Label21 = "Bandeja de entrada actualizada, " & Now & " "
Label21.ForeColor = &H0&

CommandButton9.Enabled = False
CommandButton9.BackColor = &HE0E0E0

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close


End Sub



Private Sub CommandButton12_Click()
Image13_Click
End Sub

Private Sub CommandButton13_Click()

'On Error GoTo error_Handler

ListBox2.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox2.AddItem Rs!ID_REGISTRO
Me.ListBox2.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox2.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox2.List(c, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox2.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox2.List(c, 5) = Rs!FECHA_ENVIO
c = c + 1
Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
miConexion.Close


Label76.BorderColor = &H80000001
Label95.BorderColor = &H808080
Label478.BorderColor = &H808080

Label17 = ListBox2.ListCount

Label21 = "Mi inbox actualizado, " & Now & " "
Label21.ForeColor = &H0&

CommandButton14.Enabled = False
CommandButton14.BackColor = &HE0E0E0

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close


End Sub

Private Sub CommandButton14_Click()

'On Error GoTo error_Handler

Devolver_registro = MsgBox("Realmente desea devolver esta venta a la bandeja de entrada?", vbOKCancel, "TP Ventas vodafone - Gestión Back Office")
If Devolver_registro = vbOK Then

            Call BD_Principal
            Set Rv = New ADODB.Recordset
            Rv.Open "SELECT * FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Call BD_Principal
            Set Rs = New ADODB.Recordset
            Rs.Open "SELECT * FROM TB_Principal", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
            
            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & ListBox2.Text & "'"
            
            Rs.AddNew
            Rs!ID_REGISTRO = Rv!ID_REGISTRO
            Rs!FECHA_ENVIO = Rv!FECHA_ENVIO
            Rs!USUARIO_CREADOR = Rv!USUARIO_CREADOR
            Rs!ID_CLIENTE = Rv!ID_CLIENTE
            Rs!ID_ORDEN_INTERACCION = Rv!ID_ORDEN_INTERACCION
            Rs!TIPO_VENTA = Rv!TIPO_VENTA
            Rs!TIPO_ALTA = Rv!TIPO_ALTA
            Rs!PRODUCTO_CARACTERISTICA = Rv!PRODUCTO_CARACTERISTICA
            Rs!OFERTA_ASOCIADA = Rv!OFERTA_ASOCIADA
            Rs!OBSERVACIONES = Rv!OBSERVACIONES
            Rs.Update
            
            Rv.Delete
            Rv.MoveNext
            
ListBox1.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Principal ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox1.AddItem Rs!ID_REGISTRO
Me.ListBox1.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox1.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox1.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox1.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox1.List(a, 5) = Rs!FECHA_ENVIO
a = a + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing

ListBox2.Clear

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rv.EOF = False
Me.ListBox2.AddItem Rv!ID_REGISTRO
Me.ListBox2.List(c, 1) = Rv!USUARIO_CREADOR
Me.ListBox2.List(c, 2) = Rv!ID_CLIENTE
Me.ListBox2.List(c, 3) = Rv!ID_ORDEN_INTERACCION
Me.ListBox2.List(c, 4) = Rv!TIPO_VENTA
Me.ListBox2.List(c, 5) = Rv!FECHA_ENVIO
c = c + 1
Rv.MoveNext
Wend

Rv.Close
Set Rv = Nothing
miConexion.Close

Label17 = ListBox2.ListCount
Label14 = ListBox1.ListCount

CommandButton9.Enabled = False
CommandButton9.BackColor = &HE0E0E0
CommandButton14.Enabled = False
CommandButton14.BackColor = &HE0E0E0

Label21 = "Venta devuelta a la bandeja de ventas en curso"
Label21.ForeColor = &H0&
End If

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton17_Click()

'On Error GoTo error_Handler

Label80 = "Mi inbox, ventas pendientes por gestionar"
Label80.ForeColor = &H0&
Label182 = "Mi inbox, ventas pendientes por gestionar"

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de envío"

Label49 = "ID de orden"

CommandButton17.Enabled = False
CommandButton17.BackColor = &HE0E0E0

CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&
CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&

Label76.BorderColor = &H80000001
Label95.BorderColor = &H808080
Label478.BorderColor = &H808080


TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

ListBox3.Clear

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_ENVIO,USUARIO_GESTOR FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rv.EOF = False
Me.ListBox3.AddItem Rv!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rv!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rv!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rv!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rv!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rv!FECHA_ENVIO
c = c + 1
Rv.MoveNext
Wend

Rv.Close
Set Rv = Nothing
miConexion.Close

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close


End Sub

Private Sub CommandButton18_Click()

'On Error GoTo error_Handler

Cancelar = MsgBox("Realmente desea cancelar la operación?", vbOKCancel, "TP Ventas vodafone - Gestión Back Office")
If Cancelar = vbOK Then


If Label80 = "Venta nueva, Gestión en curso..." Then
Label80 = "Mi inbox, ventas pendientes por gestionar"
End If

If Label80 = "Venta agendada, Gestión en curso..." Then
Label80 = "Mi inbox, ventas agendadas pendientes por gestionar"
End If

If Label80 = "Venta en curso, Gestión en curso..." Then
Label80 = "Mi inbox, ventas en curso pendientes por gestionar"
End If


If Label182 = "Mi inbox, ventas en curso pendientes por gestionar" Then

ListBox3.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox3.AddItem Rs!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox3.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha programada"

Label49 = "ID de orden"

CommandButton22.Enabled = False
CommandButton22.BackColor = &HE0E0E0

CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&

MultiPage2.Value = 0
MultiPage3.Value = 0

End If


If Label182 = "Mi inbox, ventas pendientes por gestionar" Then

ListBox3.Clear

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_ENVIO,USUARIO_GESTOR FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rv.EOF = False
Me.ListBox3.AddItem Rv!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rv!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rv!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rv!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rv!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rv!FECHA_ENVIO
c = c + 1
Rv.MoveNext
Wend

Rv.Close
Set Rv = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de envío"

Label49 = "ID de orden"

CommandButton17.Enabled = False
CommandButton17.BackColor = &HE0E0E0
CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&
CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&

MultiPage2.Value = 0
MultiPage3.Value = 0

End If

'_______________________________________________________________________________

If Label182 = "Mi inbox, ventas agendadas pendientes por gestionar" Then

ListBox3.Clear

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO,USUARIO_GESTOR FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rq.EOF = False
Me.ListBox3.AddItem Rq!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rq!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rq!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rq!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rq!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rq!FECHA_AGENDADO
c = c + 1
Rq.MoveNext
Wend

Rq.Close
Set Rq = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de agendado"

Label49 = "ID de orden"

CommandButton23.Enabled = False
CommandButton23.BackColor = &HE0E0E0
CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&

End If
End If

MultiPage2.Value = 0
MultiPage3.Value = 0

Image42.Visible = False
Image46.Visible = False
Label379.Visible = True

ListBox3.Enabled = True
ListBox3.BackColor = &H80000018

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""
TextBox5 = ""
TextBox2 = ""
TextBox6 = ""
TextBox77 = ""
ComboBox1 = ""
ComboBox2 = ""

ComboBox1.Enabled = False
ComboBox1.BackColor = &HE0E0E0
ComboBox2.Enabled = False
ComboBox2.BackColor = &HE0E0E0
TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0

TextBox01.Enabled = False
TextBox02.Enabled = False
TextBox03.Enabled = False
TextBox04.Enabled = False
TextBox05.Enabled = False
TextBox06.Enabled = False
TextBox07.Enabled = False
TextBox08.Enabled = False
TextBox09.Enabled = False
Label58.Enabled = False

CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0
CommandButton21.Enabled = True
CommandButton21.BackColor = &HC0&

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton19_Click()  '-------------> Terminar y enviar!

'On Error GoTo error_Handler

If ComboBox1.Text = "Venta en curso" Then
If Not IsDate(TextBox95.Text) Then
TextBox95 = ""
TextBox95 = Now
TextBox95.BackColor = &HC0C0FF
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox95.SetFocus
Exit Sub
End If

If Len(TextBox95.Text) < 18 Then
TextBox95 = ""
TextBox95 = Now
TextBox95.BackColor = &HC0C0FF
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox95.SetFocus
Exit Sub
End If
End If

If ComboBox1.Text = "Agendar venta" Then
If Not IsDate(TextBox6.Text) Then
TextBox6 = ""
TextBox6 = Now
TextBox6.BackColor = &HC0C0FF
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.SetFocus
Exit Sub
End If

If Len(TextBox6.Text) < 18 Then
TextBox6 = ""
TextBox6 = Now
TextBox6.BackColor = &HC0C0FF
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.SetFocus
Exit Sub
End If
End If

'_______________________________________________________________________________________________________________________________

Base_Datos = MsgBox("Esta seguro de terminar y enviar esta información?", vbOKCancel, "TP Ventas vodafone - Gestión Back Office")
If Base_Datos = vbOK Then

Image42.Visible = False
Image46.Visible = False
TextBox95.Visible = False
Label477.Visible = False

CommandButton19.Caption = "Enviando datos..."
CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
Application.Wait (Now + TimeValue("00:00:01"))

Label58 = Label58 & Chr(10) & "Fecha: " & Now & ""
Label58 = Label58 & Chr(10) & "Gestión Back Office: " & Label228.Caption & " - " & Label22.Caption & ""
Label58 = Label58 & Chr(10) & "Resultado: " & ComboBox1.Text & ""
Label58 = Label58 & Chr(10) & "Resultado final: " & ComboBox2.Text & ""
Label58 = Label58 & Chr(10) & "Producto: " & TextBox05.Text & "/" & TextBox06.Text & ""
Label58 = Label58 & Chr(10) & "Observaciones: " & TextBox2.Text & ""
Label58 = Label58 & Chr(10) & "**************************************************"

TextBox84.Enabled = False

If CheckBox1.Value = False Then
TextBox84 = "#"
TextBox84 = TextBox84 & Chr(10) & "**************** GESTIÓN B.O - TP VENTAS VODAFONE *****************"
TextBox84 = TextBox84 & Chr(10) & "Fecha & Hora: " & Now & ""
TextBox84 = TextBox84 & Chr(10) & "Observaciones."
TextBox84 = TextBox84 & Chr(10) & "" & TextBox2.Text & ""
TextBox84 = TextBox84 & Chr(10) & ""
TextBox84 = TextBox84 & Chr(10) & "ID del cliente: " & TextBox02.Text & ""
TextBox84 = TextBox84 & Chr(10) & "ID de la orden: " & TextBox04.Text & ""
TextBox84 = TextBox84 & Chr(10) & "Resultado de la gestión: " & ComboBox1.Text & ""
TextBox84 = TextBox84 & Chr(10) & "Resultado final: " & ComboBox2.Text & ""
TextBox84 = TextBox84 & Chr(10) & "Producto: " & TextBox05.Text & "/" & TextBox06.Text & ""
TextBox84 = TextBox84 & Chr(10) & "Usuario creador de la venta: " & TextBox08.Text & ""
TextBox84 = TextBox84 & Chr(10) & "*******************************************************************"
End If

If Label80 = "Venta nueva, Gestión en curso..." Then

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT * FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

If ComboBox1.Text = "Venta en curso" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!ID_ORDEN_BO = TextBox5.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_GESTION = TextBox95.Text
            Rg.Update
            
            Rv.Delete
            Rv.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente a ventas en curso"
            Label80.ForeColor = &H0&
End If



If ComboBox1.Text = "Agendar venta" Then
If ComboBox3.Text = "Enviar a mi inbox" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
 
            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_AGENDADO = TextBox6.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg.Update
            
            Rv.Delete
            Rv.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "La venta a sido agendada correctamente en su inbox"
            Label80.ForeColor = &H0&
            
End If

If ComboBox3.Text = "Enviar a la bandeja" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_AGENDADO = TextBox6.Text
            Rg.Update
            
            Rv.Delete
            Rv.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "La venta a sido agendada correctamente"
            Label80.ForeColor = &H0&
End If
End If

If ComboBox1.Text = "Venta exitosa" Then

            Call BD_Gestionados
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionados_EX", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg!USUARIO_CREADOR = TextBox08
            Rg!RESULTADO = ComboBox1.Text
            Rg!RESULTADO_FINAL = ComboBox2.Text
            Rg!ID_ORDEN = TextBox5.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg.Update
            
            Rv.Delete
            Rv.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente como exitosa"
            Label80.ForeColor = &H0&
End If

If ComboBox1.Text = "Venta fallida" Then

            Call BD_Gestionados
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionados_FL", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg!USUARIO_CREADOR = TextBox08
            Rg!RESULTADO = ComboBox1.Text
            Rg!RESULTADO_FINAL = ComboBox2.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox03.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg.Update
            
            Rv.Delete
            Rv.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente como fallida"
            Label80.ForeColor = &H0&
End If

'________________________ Venta nueva - Mi inbox _____________________________

Fin_Gestión = Time
Total_Gestión = Format((CDate(Inicio_Gestión) - CDate(Fin_Gestión)), "hh:mm:ss")

Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT * FROM TB_Productividad", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Pr.AddNew
Pr!ID_REGISTRO = TextBox09.Text
Pr!USUARIO_GESTOR = Label78.Caption
Pr!ID_CLIENTE = TextBox02.Text
Pr!RESULTADO = ComboBox1.Text
Pr!RESULTADO_FINAL = ComboBox2.Text
Pr!TIEMPO_GESTION = Total_Gestión
Pr.Update
             
Pr.Close
Set Pr = Nothing

ListBox3.Enabled = True
ListBox3.BackColor = &H80000018

CommandButton19.Caption = "Actualizando...."
Application.Wait (Now + TimeValue("00:00:01"))

ListBox3.Clear

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_ENVIO,USUARIO_GESTOR FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rv.EOF = False
Me.ListBox3.AddItem Rv!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rv!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rv!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rv!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rv!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rv!FECHA_ENVIO
c = c + 1
Rv.MoveNext
Wend

Rv.Close
Set Rv = Nothing

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label76 = Ra("CONTAR_REGISTROS")

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label95 = Ra("CONTAR_REGISTROS")

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label478 = Ra("CONTAR_REGISTROS")

Ra.Close
Set Ra = Nothing
miConexion.Close

CommandButton17.Enabled = False
CommandButton17.BackColor = &HE0E0E0
CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&
CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de envío"

End If

'______________________________ #Venta agendada ___________________________________

If Label80 = "Venta agendada, Gestión en curso..." Then

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

If ComboBox1.Text = "Venta en curso" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rq.MoveFirst
            Rq.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!ID_ORDEN_BO = TextBox5.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_GESTION = TextBox95.Text
            Rg.Update
            
            Rq.Delete
            Rq.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente a ventas en curso"
            Label80.ForeColor = &H0&
End If


If ComboBox1.Text = "Agendar venta" Then
If ComboBox3.Text = "Enviar a mi inbox" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rq.MoveFirst
            Rq.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_AGENDADO = TextBox6.Text
            Rg.Update
            
            Rq.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "La venta a sido agendada correctamente en su inbox"
            Label80.ForeColor = &H0&
            
End If
If ComboBox3.Text = "Enviar a la bandeja" Then
            
            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rq.MoveFirst
            Rq.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_AGENDADO = TextBox6.Text
            Rg.Update
            
            Rq.Delete
            Rq.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "La venta a sido agendada correctamente"
            Label80.ForeColor = &H0&
End If
End If

If ComboBox1.Text = "Venta exitosa" Then
       
            Call BD_Gestionados
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionados_EX", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rq.MoveFirst
            Rq.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg!USUARIO_CREADOR = TextBox08
            Rg!RESULTADO = ComboBox1.Text
            Rg!RESULTADO_FINAL = ComboBox2.Text
            Rg!ID_ORDEN = TextBox5.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg.Update
            
            Rq.Delete
            Rq.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente como exitosa"
            Label80.ForeColor = &H0&
End If

If ComboBox1.Text = "Venta fallida" Then

            Call BD_Gestionados
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionados_FL", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rq.MoveFirst
            Rq.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg!USUARIO_CREADOR = TextBox08
            Rg!RESULTADO = ComboBox1.Text
            Rg!RESULTADO_FINAL = ComboBox2.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox03.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg.Update
            
            Rq.Delete
            Rq.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente como fallida"
            Label80.ForeColor = &H0&
End If

'__________________________________ Venta agendada ________________________________________

Fin_Gestión = Time
Total_Gestión = Format((CDate(Inicio_Gestión) - CDate(Fin_Gestión)), "hh:mm:ss")

Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT * FROM TB_Productividad", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Pr.AddNew
Pr!ID_REGISTRO = TextBox09.Text
Pr!USUARIO_GESTOR = Label78.Caption
Pr!ID_CLIENTE = TextBox02.Text
Pr!RESULTADO = ComboBox1.Text
Pr!RESULTADO_FINAL = ComboBox2.Text
Pr!TIEMPO_GESTION = Total_Gestión
Pr.Update
             
Pr.Close
Set Pr = Nothing

ListBox3.Enabled = True
ListBox3.BackColor = &H80000018


CommandButton19.Caption = "Actualizando...."
Application.Wait (Now + TimeValue("00:00:01"))

ListBox3.Clear

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO,USUARIO_GESTOR FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rq.EOF = False
Me.ListBox3.AddItem Rq!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rq!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rq!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rq!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rq!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rq!FECHA_AGENDADO
c = c + 1
Rq.MoveNext
Wend

Rq.Close
Set Rq = Nothing

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label76 = Ra("CONTAR_REGISTROS")

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label95 = Ra("CONTAR_REGISTROS")

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label478 = Ra("CONTAR_REGISTROS")

Ra.Close
Set Ra = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de agendado"

CommandButton23.Enabled = False
CommandButton23.BackColor = &HE0E0E0
CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&


End If

'______________________________ #Venta en curso ___________________________________


If Label80 = "Venta en curso, Gestión en curso..." Then

Call BD_Principal
Set Rc = New ADODB.Recordset
Rc.Open "SELECT * FROM TB_Gestionando_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

If ComboBox1.Text = "Venta en curso" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rc.MoveFirst
            Rc.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!ID_ORDEN_BO = TextBox5.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_GESTION = TextBox95.Text
            Rg.Update
            
            Rc.Delete
            Rc.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente a ventas en curso"
            Label80.ForeColor = &H0&
End If

If ComboBox1.Text = "Agendar venta" Then
If ComboBox3.Text = "Enviar a mi inbox" Then
      
            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
            
            Rc.MoveFirst
            Rc.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_AGENDADO = TextBox6.Text
            Rg!USUARIO_CREADOR = Label78.Caption
            Rg.Update
            
            Rc.Delete
            Rc.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "La venta a sido agendada correctamente en su inbox"
            Label80.ForeColor = &H0&
            
End If
If ComboBox3.Text = "Enviar a la bandeja" Then

            Call BD_Principal
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rc.MoveFirst
            Rc.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!USUARIO_CREADOR = TextBox08
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg!FECHA_AGENDADO = TextBox6.Text
            Rg.Update
            
            Rc.Delete
            Rc.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "La venta a sido agendada correctamente"
            Label80.ForeColor = &H0&
End If
End If

If ComboBox1.Text = "Venta exitosa" Then

            Call BD_Gestionados
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionados_EX", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rc.MoveFirst
            Rc.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg!USUARIO_CREADOR = TextBox08
            Rg!RESULTADO = ComboBox1.Text
            Rg!RESULTADO_FINAL = ComboBox2.Text
            Rg!ID_ORDEN = TextBox5.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg.Update
            
            Rc.Delete
            Rc.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente como exitosa"
            Label80.ForeColor = &H0&
End If

If ComboBox1.Text = "Venta fallida" Then

            Call BD_Gestionados
            Set Rg = New ADODB.Recordset
            Rg.Open "SELECT * FROM TB_Gestionados_FL", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rc.MoveFirst
            Rc.Find " ID_REGISTRO = '" & TextBox09.Text & "'"

            Rg.AddNew
            Rg!ID_REGISTRO = TextBox09.Text
            Rg!USUARIO_GESTOR = Label78.Caption
            Rg!USUARIO_CREADOR = TextBox08
            Rg!RESULTADO = ComboBox1.Text
            Rg!RESULTADO_FINAL = ComboBox2.Text
            Rg!FECHA_ENVIO = TextBox01.Text
            Rg!ID_CLIENTE = TextBox02.Text
            Rg!ID_ORDEN_INTERACCION = TextBox03.Text
            Rg!TIPO_VENTA = TextBox04.Text
            Rg!TIPO_ALTA = TextBox05.Text
            Rg!PRODUCTO_CARACTERISTICA = TextBox06.Text
            Rg!OFERTA_ASOCIADA = TextBox07.Text
            Rg!OBSERVACIONES = Label58.Caption
            Rg.Update
            
            Rc.Delete
            Rc.MoveNext

            Rg.Close
            Set Rg = Nothing
 
            Label80 = "Venta enviada correctamente como fallida"
            Label80.ForeColor = &H0&
End If


'__________________________________ Venta en curso ________________________________________

Fin_Gestión = Time
Total_Gestión = Format((CDate(Inicio_Gestión) - CDate(Fin_Gestión)), "hh:mm:ss")

Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT * FROM TB_Productividad", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Pr.AddNew
Pr!ID_REGISTRO = TextBox09.Text
Pr!USUARIO_GESTOR = Label78.Caption
Pr!ID_CLIENTE = TextBox02.Text
Pr!RESULTADO = ComboBox1.Text
Pr!RESULTADO_FINAL = ComboBox2.Text
Pr!TIEMPO_GESTION = Total_Gestión
Pr.Update
             
Pr.Close
Set Pr = Nothing

CommandButton19.Caption = "Actualizando...."
Application.Wait (Now + TimeValue("00:00:01"))

ListBox3.Enabled = True
ListBox3.BackColor = &H80000018

ListBox3.Clear

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rq.EOF = False
Me.ListBox3.AddItem Rq!ID_REGISTRO
Me.ListBox3.List(X, 1) = Rq!USUARIO_CREADOR
Me.ListBox3.List(X, 2) = Rq!ID_CLIENTE
Me.ListBox3.List(X, 3) = Rq!ID_ORDEN_BO
Me.ListBox3.List(X, 4) = Rq!TIPO_VENTA
Me.ListBox3.List(X, 5) = Rq!FECHA_GESTION
X = X + 1
Rq.MoveNext
Wend

Rq.Close
Set Rq = Nothing


Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label76 = Ra("CONTAR_REGISTROS")

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label95 = Ra("CONTAR_REGISTROS")

Call BD_Principal
Set Ra = New ADODB.Recordset
Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label478 = Ra("CONTAR_REGISTROS")

Ra.Close
Set Ra = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha programada"

CommandButton22.Enabled = False
CommandButton22.BackColor = &HE0E0E0
CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&

End If

'____________________________________ Ejecución normal de codigo final __________________________________


MultiPage3.Value = 0
MultiPage2.Value = 0
Label379.Visible = True

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

TextBox01.Enabled = False
TextBox02.Enabled = False
TextBox03.Enabled = False
TextBox04.Enabled = False
TextBox05.Enabled = False
TextBox06.Enabled = False
TextBox07.Enabled = False
TextBox08.Enabled = False
TextBox09.Enabled = False
Label58.Enabled = False

ComboBox1 = ""
ComboBox2 = ""
ComboBox3 = ""
TextBox5 = ""
TextBox2 = ""
TextBox6 = ""


CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
CommandButton21.Enabled = True
CommandButton21.BackColor = &HC0&

ComboBox1.Enabled = False
ComboBox1.BackColor = &HE0E0E0
ComboBox2.Enabled = False
ComboBox2.BackColor = &HE0E0E0
ComboBox3.Enabled = False
ComboBox3.BackColor = &HE0E0E0
TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0

If CheckBox1.Value = False Then
MultiPage1.Value = 9
TextBox84.Enabled = True
End If

CommandButton19.Caption = "Terminar y enviar"
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0

End If

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton21_Click()

Image11_Click

End Sub

Private Sub CommandButton22_Click()

Label80 = "Mi inbox, ventas en curso pendientes por gestionar"
Label80.ForeColor = &H0&
Label182 = "Mi inbox, ventas en curso pendientes por gestionar"

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha programada"

Label49 = "ID de orden"

CommandButton22.Enabled = False
CommandButton22.BackColor = &HE0E0E0

CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&

Label76.BorderColor = &H808080
Label95.BorderColor = &H808080
Label478.BorderColor = &H80000001

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

ListBox3.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox3.AddItem Rs!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox3.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
miConexion.Close

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close


End Sub

Private Sub CommandButton23_Click()


'On Error GoTo error_Handler

Label80 = "Mi inbox, ventas agendadas pendientes por gestionar"
Label80.ForeColor = &H0&

Label182 = "Mi inbox, ventas agendadas pendientes por gestionar"

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de agendado"

Label49 = "ID de orden"

CommandButton23.Enabled = False
CommandButton23.BackColor = &HE0E0E0

CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&

Label76.BorderColor = &H808080
Label95.BorderColor = &H80000001
Label478.BorderColor = &H808080

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

ListBox3.Clear

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO,USUARIO_GESTOR FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rq.EOF = False
Me.ListBox3.AddItem Rq!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rq!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rq!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rq!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rq!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rq!FECHA_AGENDADO
c = c + 1
Rq.MoveNext
Wend

Rq.Close
Set Rq = Nothing
miConexion.Close

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close
End Sub

Private Sub CommandButton24_Click()
MultiPage2.Value = 0
TextBox77 = ""
End Sub


Private Sub CommandButton32_Click()

'On Error GoTo error_Handler


ListBox4.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Agendado ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox4.AddItem Rs!ID_REGISTRO
Me.ListBox4.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox4.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox4.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox4.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox4.List(a, 5) = Rs!FECHA_AGENDADO
a = a + 1
Rs.MoveNext
Loop

Label128 = ListBox4.ListCount
Label184 = Now
Label136 = "Ventas agendadas actualizada, " & Label184.Caption & " "
Label136.ForeColor = &H0&

Rs.Close
Set Rs = Nothing
miConexion.Close

CommandButton33.Enabled = False
CommandButton33.BackColor = &HE0E0E0

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton33_Click()

On Error GoTo Error_obtener

If Label142 >= "5" Then
Label136.ForeColor = &HC0&
Label136 = "Superó el límite de ventas agendadas que puede obtener"
Exit Sub
Else

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


            Rs.MoveFirst
            Rs.Find " ID_REGISTRO = '" & ListBox4.Text & "'"
            
            Rq.AddNew
            Rq!ID_REGISTRO = Rs!ID_REGISTRO
            Rq!FECHA_ENVIO = Rs!FECHA_ENVIO
            Rq!USUARIO_CREADOR = Rs!USUARIO_CREADOR
            Rq!ID_CLIENTE = Rs!ID_CLIENTE
            Rq!ID_ORDEN_INTERACCION = Rs!ID_ORDEN_INTERACCION
            Rq!TIPO_VENTA = Rs!TIPO_VENTA
            Rq!TIPO_ALTA = Rs!TIPO_ALTA
            Rq!PRODUCTO_CARACTERISTICA = Rs!PRODUCTO_CARACTERISTICA
            Rq!OFERTA_ASOCIADA = Rs!OFERTA_ASOCIADA
            Rq!OBSERVACIONES = Rs!OBSERVACIONES
            Rq!FECHA_AGENDADO = Rs!FECHA_AGENDADO
            Rq!USUARIO_GESTOR = Label22.Caption
            Rq.Update
            
            Rs.Delete
            Rs.MoveNext
 
Rq.Close
Set Rq = Nothing


ListBox4.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Agendado ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox4.AddItem Rs!ID_REGISTRO
Me.ListBox4.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox4.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox4.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox4.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox4.List(a, 5) = Rs!FECHA_AGENDADO
a = a + 1
Rs.MoveNext
Loop

ListBox5.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox5.AddItem Rs!ID_REGISTRO
Me.ListBox5.List(i, 1) = Rs!USUARIO_CREADOR
Me.ListBox5.List(i, 2) = Rs!ID_CLIENTE
Me.ListBox5.List(i, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox5.List(i, 4) = Rs!TIPO_VENTA
Me.ListBox5.List(i, 5) = Rs!FECHA_AGENDADO
i = i + 1
Rs.MoveNext
Loop

Label128 = ListBox4.ListCount
Label142 = ListBox5.ListCount

Rs.Close
Set Rs = Nothing
miConexion.Close

CommandButton33.Enabled = False
CommandButton33.BackColor = &HE0E0E0

Label136 = "Venta obtenida correctamente en su inbox"
Label136.ForeColor = &H0&

End If

Exit Sub
            
'____________________________________ Error obtener venta _______________________________________________________________
            
Error_obtener:

On Error GoTo Error_obtener2
                
Label136 = "La venta fue obtenida por otro usuario. Por favor, inténtalo de nuevo"
Label136.ForeColor = &HC0&

ListBox4.Clear

Call BD_Principal
Set Re = New ADODB.Recordset
Re.Open "SELECT * FROM TB_Agendado ORDER BY REANGEDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Re.EOF = False
Me.ListBox4.AddItem Re!FECHA_Y_HORA
Me.ListBox4.List(G, 1) = Re!NOMBRE_Y_APELLIDO
Me.ListBox4.List(G, 2) = Re!NUMERO_DOCUMENTO
Me.ListBox4.List(G, 3) = Re!MOVIL_CONTACTO
Me.ListBox4.List(G, 4) = Re!PRODUCTO_VENTA
Me.ListBox4.List(G, 5) = Re!REANGEDADO
G = G + 1
Re.MoveNext
Wend

Re.Close
Set Re = Nothing

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rq.Close
Set Rq = Nothing
miConexion.Close

Label128 = ListBox4.ListCount

CommandButton9.Enabled = False
CommandButton9.BackColor = &HE0E0E0

Exit Sub
Error_obtener2:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton35_Click()

'On Error GoTo error_Handler

ListBox3.Clear

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO,USUARIO_GESTOR FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rq.EOF = False
Me.ListBox3.AddItem Rq!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rq!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rq!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rq!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rq!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rq!FECHA_AGENDADO
c = c + 1
Rq.MoveNext
Wend

Rq.Close
Set Rq = Nothing

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT COUNT(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label76 = Rs("CONTAR_REGISTROS")

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label95 = Rs("CONTAR_REGISTROS")

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label478 = Rs("CONTAR_REGISTROS")

Rs.Close
Set Rs = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de agendado"

Label49 = "ID de orden"

Label76.BorderColor = &H808080
Label95.BorderColor = &H80000001
Label478.BorderColor = &H808080

CommandButton23.Enabled = False
CommandButton23.BackColor = &HE0E0E0

CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&
CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&

'________________________________________________-

CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0

ComboBox1.Enabled = False
ComboBox1.BackColor = &HE0E0E0
ComboBox2.Enabled = False
ComboBox2.BackColor = &HE0E0E0
ComboBox3.Enabled = False
ComboBox3.BackColor = &HE0E0E0

TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

TextBox01.Enabled = False
TextBox02.Enabled = False
TextBox03.Enabled = False
TextBox04.Enabled = False
TextBox05.Enabled = False
TextBox06.Enabled = False
TextBox07.Enabled = False
TextBox08.Enabled = False
TextBox09.Enabled = False
Label58.Enabled = False


MultiPage1.Value = 1
MultiPage2.Value = 0
MultiPage3.Value = 0
UserForm1.Caption = "TP Ventas vodafone - Mi inbox"

Label80 = "Mi inbox, ventas agendadas pendientes por gestionar"
Label80.ForeColor = &H0&

Label182 = "Mi inbox, ventas agendadas pendientes por gestionar"

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton36_Click()

'On Error GoTo error_Handler

ListBox5.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox5.AddItem Rs!ID_REGISTRO
Me.ListBox5.List(i, 1) = Rs!USUARIO_CREADOR
Me.ListBox5.List(i, 2) = Rs!ID_CLIENTE
Me.ListBox5.List(i, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox5.List(i, 4) = Rs!TIPO_VENTA
Me.ListBox5.List(i, 5) = Rs!FECHA_AGENDADO
i = i + 1
Rs.MoveNext
Loop
Label142 = ListBox5.ListCount

Rs.Close
Set Rs = Nothing
miConexion.Close

Label184 = Now
Label136 = "Mi inbox actualizado, " & Label184.Caption & " "
Label136.ForeColor = &H0&


CommandButton37.Enabled = False
CommandButton37.BackColor = &HE0E0E0

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton37_Click()

'On Error GoTo error_Handler

Devolver_registro = MsgBox("Realmente desea devolver esta venta a la bandeja de ventas agendadas?", vbOKCancel, "TP Ventas vodafone - Gestión Back Office")
If Devolver_registro = vbOK Then


Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Principal
Set Rq = New ADODB.Recordset
Rq.Open "SELECT * FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

            Rs.MoveFirst
            Rs.Find " ID_REGISTRO = '" & ListBox5.Text & "'"
            
            Rq.AddNew
            Rq!ID_REGISTRO = Rs!ID_REGISTRO
            Rq!FECHA_ENVIO = Rs!FECHA_ENVIO
            Rq!USUARIO_CREADOR = Rs!USUARIO_CREADOR
            Rq!ID_CLIENTE = Rs!ID_CLIENTE
            Rq!ID_ORDEN_INTERACCION = Rs!ID_ORDEN_INTERACCION
            Rq!TIPO_VENTA = Rs!TIPO_VENTA
            Rq!TIPO_ALTA = Rs!TIPO_ALTA
            Rq!PRODUCTO_CARACTERISTICA = Rs!PRODUCTO_CARACTERISTICA
            Rq!OFERTA_ASOCIADA = Rs!OFERTA_ASOCIADA
            Rq!OBSERVACIONES = Rs!OBSERVACIONES
            Rq!FECHA_AGENDADO = Rs!FECHA_AGENDADO
            Rq.Update
            
            Rs.Delete
            Rs.MoveNext
 
Rq.Close
Set Rq = Nothing


ListBox4.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Agendado ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox4.AddItem Rs!ID_REGISTRO
Me.ListBox4.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox4.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox4.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox4.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox4.List(a, 5) = Rs!FECHA_AGENDADO
a = a + 1
Rs.MoveNext
Loop

ListBox5.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox5.AddItem Rs!ID_REGISTRO
Me.ListBox5.List(i, 1) = Rs!USUARIO_CREADOR
Me.ListBox5.List(i, 2) = Rs!ID_CLIENTE
Me.ListBox5.List(i, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox5.List(i, 4) = Rs!TIPO_VENTA
Me.ListBox5.List(i, 5) = Rs!FECHA_AGENDADO
i = i + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing
miConexion.Close


Label128 = ListBox4.ListCount
Label142 = ListBox5.ListCount

CommandButton37.Enabled = False
CommandButton37.BackColor = &HE0E0E0

Label136 = "Venta devuelta a la bandeja de ventas agendadas"
Label136.ForeColor = &H0&

End If

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton39_Click()
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True
End Sub

Private Sub CommandButton40_Click()

Label195 = "ID de orden"

ListBox6.Clear

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,RESULTADO,RESULTADO_FINAL,FECHA_FIN_GESTION,USUARIO_GESTOR FROM TB_Gestionados_FL WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
While Rg.EOF = False

Me.ListBox6.AddItem Rg!ID_REGISTRO
Me.ListBox6.List(Z, 1) = Rg!USUARIO_CREADOR
Me.ListBox6.List(Z, 2) = Rg!ID_CLIENTE
Me.ListBox6.List(Z, 3) = Rg!ID_ORDEN_INTERACCION
Me.ListBox6.List(Z, 4) = Rg!RESULTADO
Me.ListBox6.List(Z, 5) = Rg!RESULTADO_FINAL
Me.ListBox6.List(Z, 6) = Rg!FECHA_FIN_GESTION
Z = Z + 1
Rg.MoveNext
Wend

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionados_EX WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label201 = Rg("CONTAR_REGISTROS")

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionados_FL WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label191 = Rg("CONTAR_REGISTROS")

Label205.Caption = Val(Label201.Caption) + Val(Label191.Caption)

Rg.Close
Set Rg = Nothing
miConexion.Close

CommandButton41.Enabled = True
CommandButton41.BackColor = &HC0&
CommandButton40.Enabled = False
CommandButton40.BackColor = &HE0E0E0

Label199 = "Mis ventas gestionadas con resultado fallido"
Label199.ForeColor = &H0&


End Sub

Private Sub CommandButton41_Click()

Label195 = "ID de orden"

ListBox6.Clear

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN,RESULTADO,RESULTADO_FINAL,FECHA_FIN_GESTION,USUARIO_GESTOR FROM TB_Gestionados_EX WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
While Rg.EOF = False

Me.ListBox6.AddItem Rg!ID_REGISTRO
Me.ListBox6.List(Z, 1) = Rg!USUARIO_CREADOR
Me.ListBox6.List(Z, 2) = Rg!ID_CLIENTE
Me.ListBox6.List(Z, 3) = Rg!ID_ORDEN
Me.ListBox6.List(Z, 4) = Rg!RESULTADO
Me.ListBox6.List(Z, 5) = Rg!RESULTADO_FINAL
Me.ListBox6.List(Z, 6) = Rg!FECHA_FIN_GESTION
Z = Z + 1
Rg.MoveNext
Wend

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionados_EX WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label201 = Rg("CONTAR_REGISTROS")

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionados_FL WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label191 = Rg("CONTAR_REGISTROS")

Label205.Caption = Val(Label201.Caption) + Val(Label191.Caption)

Rg.Close
Set Rg = Nothing
miConexion.Close

CommandButton40.Enabled = True
CommandButton40.BackColor = &HC0&
CommandButton41.Enabled = False
CommandButton41.BackColor = &HE0E0E0

Label199 = "Mis ventas gestionadas como exitosas"
Label199.ForeColor = &H0&

End Sub



Private Sub CommandButton46_Click()

''On Error GoTo error_Handler

If txtUsuario.Text = "" Or txtPassword.Text = "" Then
MsgBox "Para validar su identidad debe ingresar un nombre de usuario y una contraseña. ", vbInformation, "Inicio de sesión"
Exit Sub

Else

Label435.ForeColor = &HC0&
Application.Wait (Now + TimeValue("00:00:01"))

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "select * from Usuarios_BO_no_autorizados where USUARIO_CITRIX ='" + txtUsuario.Text + "' and CONTRASEÑA='" + txtPassword.Text + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
If Rn.EOF = False Then

Rn.Close
Set Rn = Nothing
miConexion.Close
Label435.ForeColor = &HFFFFFF
MsgBox "Este usuario no se encuentra autorizado para gestionar ventas, por favor contacte al administrador.", vbCritical, "Inicio de sesión"
txtPassword = ""
txtPassword.SetFocus

Else

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "select * from Usuarios_BO_autorizados where USUARIO_CITRIX ='" + txtUsuario.Text + "' and CONTRASEÑA='" + txtPassword.Text + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
If Ry.EOF = True Then

Rn.Close
Set Rn = Nothing
Ry.Close
Set Ry = Nothing
miConexion.Close
Label435.ForeColor = &HFFFFFF
MsgBox "El nombre de usuario y la contraseña que ingresaste no coinciden con nuestros registros. Por favor, revisa e inténtalo de nuevo.", vbCritical, "Inicio de sesión"
txtPassword = ""
txtPassword.SetFocus


Else

    '___________________________________________ Inicio del programa ______________________________________________

    Dim contarinbox(1 To 3) As Integer

    Call BD_Seguridad
    Set Ry = New ADODB.Recordset
    Ry.Open "select * from Usuarios_BO_autorizados where USUARIO_CITRIX ='" + txtUsuario.Text + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    Ry.Find "USUARIO_CITRIX = '" & txtUsuario.Text & "'", , , 1
    If Ry.BOF = False And Ry.EOF = False Then
    Label228.Caption = Ry.Fields("NOMBRE_Y_APELLIDO")
    Label377.Caption = Ry.Fields("FECHA_REGISTRO")
    
    Ry.Close
    Set Ry = Nothing
    
    txtUsuario = UCase(txtUsuario)
    Label22.Caption = txtUsuario.Text
    Label78.Caption = Label22.Caption
    Label140.Caption = Label22.Caption
    Label203.Caption = Label22.Caption
    Label500.Caption = Label22.Caption
    Label396.Caption = Label22.Caption

    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(USUARIO_GESTOR) AS CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    contarinbox(1) = Ra("CONTAR_REGISTROS")
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(USUARIO_GESTOR) as CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    contarinbox(2) = Ra("CONTAR_REGISTROS")
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(USUARIO_GESTOR) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    contarinbox(3) = Ra("CONTAR_REGISTROS")
    
    Label371.Caption = Val(contarinbox(1)) + Val(contarinbox(2)) + Val(contarinbox(3))
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Label370 = Ra("CONTAR_REGISTROS")

    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Principal", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Label369 = Ra("CONTAR_REGISTROS")
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Label507 = Ra("CONTAR_REGISTROS")
    

    Ra.Close
    Set Ra = Nothing
    Rn.Close
    Set Rn = Nothing
    miConexion.Close
     
    Label435.ForeColor = &HFFFFFF
    Label339.Caption = "WELCOME " & Label228.Caption
    
    UserForm1.Caption = "TP Ventas vodafone - Menú principal"
    MultiPage1.Value = 6
    

End If
End If
End If
End If


Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton47_Click()

'On Error GoTo error_Handler

If TextBox80.Text = "" Or TextBox81.Text = "" Or TextBox82.Text = "" Or TextBox83.Text = "" Or ComboBox4.Text = "" Then
MsgBox "Hay campos vacios, debe completar todos los campos para poder registrarlo. ", vbInformation, "TP Ventas vodafone - Registrarse"
Exit Sub

Else

If TextBox82.Text <> TextBox83.Text Then
MsgBox "Las contraseñas no coinciden. Vuelve a intentarlo ", vbCritical, "TP Ventas vodafone - Registrarse"
TextBox82 = ""
TextBox83 = ""
TextBox82.SetFocus
Exit Sub

Else

Call BD_Seguridad
Set Rn = New ADODB.Recordset
Rn.Open "select * from Usuarios_BO_No_autorizados where USUARIO_CITRIX ='" + TextBox81 + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
If Rn.EOF = False Then
MsgBox "El usuario: " & TextBox81.Text & " ya se encuentra registrado, pero no esta autorizado para gestionar ventas. Por favor, contacte al administrador.", vbCritical, "TP Ventas vodafone - Registrarse"
TextBox80 = ""
TextBox81 = ""
TextBox82 = ""
TextBox83 = ""
ComboBox4 = ""
TextBox80.SetFocus

Rn.Close
Set Rn = Nothing
miConexion.Close
Exit Sub
 
Else
 
Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "select * from Usuarios_BO_autorizados where USUARIO_CITRIX ='" + TextBox81 + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
If Ry.EOF = False Then
MsgBox "El usuario: " & TextBox81.Text & " ya se encuentra registrado, y esta autorizado para gestionar ventas. Si presenta problemas para iniciar sesión contacte al administrador.", vbCritical, "TP Ventas vodafone - Registrarse"
TextBox80 = ""
TextBox81 = ""
TextBox82 = ""
TextBox83 = ""
ComboBox4 = ""
TextBox80.SetFocus

Ry.Close
Set Ry = Nothing
Rn.Close
Set Rn = Nothing
miConexion.Close

Exit Sub
 
Else
 
 
MsgBox "Aviso importante: Una vez guardada esta información no podra modificarla, si presenta problemas para iniciar sesión debe contactar al administrador.", vbExclamation, "TP Ventas vodafone - Registrarse"
 
Base_Datos = MsgBox("Esta seguro de terminar y enviar esta información?", vbOKCancel, "TP Ventas vodafone - Registrarse")
If Base_Datos = vbOK Then


TextBox80 = UCase(TextBox80)
TextBox81 = UCase(TextBox81)
TextBox82 = UCase(TextBox82)
  
   '_____________________________________________________
   
    Call BD_Seguridad
    Set Rn = New ADODB.Recordset
    Rn.Open "SELECT * FROM Usuarios_BO_No_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    Rn.AddNew
    Rn.Fields("NOMBRE_Y_APELLIDO") = TextBox80.Text
    Rn.Fields("USUARIO_CITRIX") = TextBox81.Text
    Rn.Fields("CONTRASEÑA") = TextBox82.Text
    Rn.Fields("CAMPAÑA") = ComboBox4.Text
    Rn.Update
    
Rn.Close
Set Rn = Nothing
Ry.Close
Set Ry = Nothing
miConexion.Close
 '_____________________________________________________
    
MsgBox "Usuario registrado correctamente, recuerde que el administrador debe autorizar el acceso.", vbInformation, "TP Ventas vodafone - Registrarse"
  
TextBox80 = ""
TextBox81 = ""
TextBox82 = ""
TextBox83 = ""
ComboBox4 = ""
TextBox80.SetFocus

MultiPage4.Value = 0


End If
End If
End If
End If
End If



Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub



Private Sub CommandButton64_Click()

'On Error GoTo error_Handler

If TextBox87.Text = "" Or TextBox85.Text = "" Or TextBox86.Text = "" Then
MsgBox "Hay campos vacios, para realizar el cambio de contraseña debe completar todos los campos.", vbInformation, "TP Ventas vodafone - Cambio de contraseña"
Exit Sub

Else

If TextBox85.Text <> TextBox86.Text Then
MsgBox "Las contraseñas no coinciden. Vuelve a intentarlo ", vbCritical, "TP Ventas vodafone - Cambio de contraseña"
TextBox85 = ""
TextBox86 = ""
TextBox85.SetFocus
Exit Sub

Else

Call BD_Seguridad
Set Ru = New ADODB.Recordset
Ru.Open "select * from Usuarios_BO_autorizados where CONTRASEÑA= '" + TextBox87.Text + "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

If Ru.EOF = True Then
MsgBox "La contraseña actual que ingresastes no coincide con nuestros registros. Por favor, revisa e inténtalo de nuevo.", vbCritical, "TP Ventas vodafone - Cambio de contraseña"
TextBox87.Text = ""
TextBox85.Text = ""
TextBox86.Text = ""
TextBox87.SetFocus

Ru.Close
Set Ru = Nothing
miConexion.Close

Exit Sub

Else

Base_Datos = MsgBox("Esta seguro de terminar y enviar esta información?", vbOKCancel, "TP Ventas vodafone - Cambio de contraseña")
If Base_Datos = vbOK Then

Call BD_Seguridad
Set Ry = New ADODB.Recordset
Ry.Open "SELECT * FROM Usuarios_BO_autorizados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Ry.MoveFirst
Ry.Find " FECHA_REGISTRO = '" & Label377.Caption & "'"
Ry!CONTRASEÑA = TextBox85.Text
Ry.Update

If Ry.State = 1 Or Ry.State = 0 Then
MsgBox "El cambio de contraseña se ha realizado correctamente.", vbInformation, "TP Ventas vodafone - Cambio de contraseña"
TextBox87.Text = ""
TextBox85.Text = ""
TextBox86.Text = ""
MultiPage1.Value = 6

Ru.Close
Set Ru = Nothing
Ry.Close
Set Ry = Nothing
miConexion.Close

Else
MsgBox "Ha ocurrido un error inesperado, no se ha podido realizar el cambio de contraseña.", vbCritical, "TP Ventas vodafone - Cambio de contraseña"
'===============================================
TextBox87.Text = ""
TextBox85.Text = ""
TextBox86.Text = ""
TextBox87.SetFocus

Ru.Close
Set Ru = Nothing
Ry.Close
Set Ry = Nothing
miConexion.Close

End If
End If
End If
End If
End If

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label81.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub CommandButton68_Click()

If Not IsDate(TextBox88.Text) Then
MsgBox "Error en la consulta, el formato de fecha no es válido", vbInformation, "TP Ventas vodafone - Panel de control"
Exit Sub

Else

If Not IsDate(TextBox89.Text) Then
MsgBox "Error en la consulta, el formato de fecha no es válido", vbInformation, "TP Ventas vodafone - Panel de control"
Exit Sub

Else

ListBox7.Clear

Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT ID_REGISTRO, USUARIO_GESTOR, ID_CLIENTE, RESULTADO,RESULTADO_FINAL, TIEMPO_GESTION, FECHA_FIN_GESTION FROM TB_Productividad WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox88.Text & "# AND #" & Label428.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Pr.EOF = False
Me.ListBox7.AddItem Pr!ID_REGISTRO
Me.ListBox7.List(e, 1) = Pr!FECHA_FIN_GESTION
Me.ListBox7.List(e, 2) = Pr!ID_CLIENTE
Me.ListBox7.List(e, 3) = Pr!RESULTADO
Me.ListBox7.List(e, 4) = Pr!RESULTADO_FINAL
Me.ListBox7.List(e, 5) = Pr!TIEMPO_GESTION
e = e + 1
Pr.MoveNext
Wend

If TextBox88.Text = TextBox89.Text Then

Label392 = "Historial de ventas gestionadas el dia de hoy " & Date & ""
Label392.ForeColor = &H0&

Else

Label392 = "Historial de ventas gestionadas entre el " & TextBox88.Text & " y " & TextBox89.Text & ""
Label392.ForeColor = &H0&

End If

Pr.Close
Set Pr = Nothing
miConexion.Close

Label527 = ListBox7.ListCount

If Label527 <> "0" Then
CommandButton70.Enabled = True
CommandButton70.BackColor = &HC0&
Else
CommandButton70.Enabled = False
CommandButton70.BackColor = &HE0E0E0
End If

End If
End If

End Sub

Private Sub CommandButton69_Click()


End Sub

Private Sub CommandButton70_Click()

Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT Count(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Productividad WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox88.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label344 = Pr("CONTAR_REGISTROS")

Pr.Close
Set Pr = Nothing
miConexion.Close

If Label344.Caption = 0 Then
MsgBox "No se encontraron datos para la exportación", vbInformation, "Sistema de exportación"
Exit Sub

Else
  
Base_Datos = MsgBox("Esta seguro de comenzar con la expotación de " & Label344.Caption & " registros?", vbOKCancel, "Sistema de exportación")
If Base_Datos = vbOK Then

CommandButton70.Enabled = False
CommandButton70.BackColor = &HE0E0E0
 
CommandButton70.Caption = "Exportando datos..."
Application.Wait (Now + TimeValue("00:00:01"))

    
Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT * FROM TB_Productividad WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox88.Text & "# AND #" & Label428.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
NombreHoja = "Productividad by " & Label22.Caption & ""

Set APIExcel = CreateObject("Excel.Application")
Set AddLibro = APIExcel.Workbooks.Add
APIExcel.Visible = False

Set AddHoja = AddLibro.Worksheets(1)
If Len(NombreHoja) > 0 Then AddHoja.Name = Left(NombreHoja, 30)
columnas = Pr.Fields.Count
For i = 0 To columnas - 1
APIExcel.Cells(1, i + 1) = Pr.Fields(i).Name
Next i

Pr.MoveFirst
AddHoja.Range("A2").CopyFromRecordset Pr

With APIExcel.ActiveSheet.Cells
.Select
.EntireColumn("A:F").AutoFit
.Columns("F").NumberFormat = "hh:mm:ss"
.Range("A1").Select
End With

Pr.Close
Set Pr = Nothing
miConexion.Close

CommandButton70.Enabled = True
CommandButton70.BackColor = &HC0&

CommandButton70.Caption = "Exportar a Microsoft Excel"

MsgBox "La expotación de " & Label344.Caption & " registros ha finalizado.", vbInformation, "Sistema de exportación"
APIExcel.Visible = True

End If
End If


End Sub

Private Sub CommandButton71_Click()

ListBox8.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_GESTION,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_BO,TIPO_VENTA FROM TB_Venta_Curso ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox8.AddItem Rs!ID_REGISTRO
Me.ListBox8.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox8.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox8.List(a, 3) = Rs!ID_ORDEN_BO
Me.ListBox8.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox8.List(a, 5) = Rs!FECHA_GESTION
a = a + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing
miConexion.Close

Label481 = ListBox8.ListCount
Label502 = "Bandeja de ventas en curso actualizada, " & Now & " "
Label502.ForeColor = &H0&

CommandButton72.Enabled = False
CommandButton72.BackColor = &HE0E0E0

End Sub

Private Sub CommandButton72_Click()

'On Error GoTo Error_obtener

If Label491 = "5" Then
Label21.ForeColor = &HC0&
Label502 = "Superó el límite de ventas en curso permitidas que puede obtener"
Exit Sub

Else

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT * FROM TB_Gestionando_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


            Rs.MoveFirst
            Rs.Find " ID_REGISTRO = '" & ListBox8.Text & "'"
            
            Rv.AddNew
            Rv!ID_REGISTRO = Rs!ID_REGISTRO
            Rv!FECHA_ENVIO = Rs!FECHA_ENVIO
            Rv!USUARIO_CREADOR = Rs!USUARIO_CREADOR
            Rv!ID_CLIENTE = Rs!ID_CLIENTE
            Rv!ID_ORDEN_BO = Rs!ID_ORDEN_BO
            Rv!TIPO_VENTA = Rs!TIPO_VENTA
            Rv!TIPO_ALTA = Rs!TIPO_ALTA
            Rv!PRODUCTO_CARACTERISTICA = Rs!PRODUCTO_CARACTERISTICA
            Rv!OFERTA_ASOCIADA = Rs!OFERTA_ASOCIADA
            Rv!OBSERVACIONES = Rs!OBSERVACIONES
            Rv!FECHA_GESTION = Rs!FECHA_GESTION
            Rv!USUARIO_GESTOR = Label22.Caption
            Rv.Update
            
            Rs.Delete
            Rs.MoveNext

Rv.Close
Set Rv = Nothing

ListBox8.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_GESTION,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_BO,TIPO_VENTA FROM TB_Venta_Curso ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox8.AddItem Rs!ID_REGISTRO
Me.ListBox8.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox8.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox8.List(a, 3) = Rs!ID_ORDEN_BO
Me.ListBox8.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox8.List(a, 5) = Rs!FECHA_GESTION
a = a + 1
Rs.MoveNext
Loop

ListBox9.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox9.AddItem Rs!ID_REGISTRO
Me.ListBox9.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox9.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox9.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox9.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox9.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS2 FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label489 = Rs("CONTAR_REGISTROS2")

Rs.Close
Set Rs = Nothing
miConexion.Close

CommandButton72.Enabled = False
CommandButton72.BackColor = &HE0E0E0

Label481 = ListBox8.ListCount
Label491 = ListBox9.ListCount

Label502 = "Venta obtenida correctamente en su inbox"
Label502.ForeColor = &H0&

End If

Exit Sub
            
'___________________________________ Error al obtener un registro ________________________________________________________________
            
Error_obtener:
On Error GoTo Error_obtener2
              
Label502 = "La venta fue obtenida por otro usuario. Por favor, inténtalo de nuevo"
Label502.ForeColor = &HC0&

ListBox8.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_GESTION,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_BO,TIPO_VENTA FROM TB_Venta_Curso ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox8.AddItem Rs!ID_REGISTRO
Me.ListBox8.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox8.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox8.List(a, 3) = Rs!ID_ORDEN_BO
Me.ListBox8.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox8.List(a, 5) = Rs!FECHA_GESTION
a = a + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT ID_REGISTRO FROM TB_Gestionando_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rv.Close
Set Rv = Nothing
miConexion.Close

Label481 = ListBox8.ListCount

CommandButton72.Enabled = False
CommandButton72.BackColor = &HE0E0E0

Exit Sub
Error_obtener2:

MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close
End Sub

Private Sub CommandButton73_Click()

Label80 = "Mi inbox, ventas en curso pendientes por gestionar"
Label80.ForeColor = &H0&
Label182 = "Mi inbox, ventas en curso pendientes por gestionar"


ListBox3.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_BO,TIPO_VENTA,FECHA_GESTION,USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox3.AddItem Rs!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox3.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT COUNT(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label76 = Rs("CONTAR_REGISTROS")

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label95 = Rs("CONTAR_REGISTROS")

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label478 = Rs("CONTAR_REGISTROS")

Rs.Close
Set Rs = Nothing
miConexion.Close


Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha programada"

Label49 = "ID de orden"

Label76.BorderColor = &H808080
Label95.BorderColor = &H808080
Label478.BorderColor = &H80000001

CommandButton22.Enabled = False
CommandButton22.BackColor = &HE0E0E0
CommandButton17.Enabled = True
CommandButton17.BackColor = &HC0&
CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&

'_____________________________________________________

CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0

ComboBox1.Enabled = False
ComboBox1.BackColor = &HE0E0E0
ComboBox2.Enabled = False
ComboBox2.BackColor = &HE0E0E0
ComboBox3.Enabled = False
ComboBox3.BackColor = &HE0E0E0

TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

TextBox01.Enabled = False
TextBox02.Enabled = False
TextBox03.Enabled = False
TextBox04.Enabled = False
TextBox05.Enabled = False
TextBox06.Enabled = False
TextBox07.Enabled = False
TextBox08.Enabled = False
TextBox09.Enabled = False
Label58.Enabled = False

MultiPage1.Value = 1
MultiPage2.Value = 0
MultiPage3.Value = 0
UserForm1.Caption = "TP Ventas vodafone - Mi inbox"

End Sub

Private Sub CommandButton74_Click()

ListBox9.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox9.AddItem Rs!ID_REGISTRO
Me.ListBox9.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox9.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox9.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox9.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox9.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
miConexion.Close

Label491 = ListBox8.ListCount
Label502 = "Mi inbox actualizado, " & Now & " "
Label502.ForeColor = &H0&

CommandButton72.Enabled = False
CommandButton72.BackColor = &HE0E0E0
End Sub

Private Sub CommandButton75_Click()


Devolver_registro = MsgBox("Realmente desea devolver esta venta a la bandeja de entrada?", vbOKCancel, "TP Ventas vodafone - Gestión Back Office")
If Devolver_registro = vbOK Then
   
Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT * FROM TB_Gestionando_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


            Rv.MoveFirst
            Rv.Find " ID_REGISTRO = '" & ListBox9.Text & "'"
            
            Rs.AddNew
            Rs!ID_REGISTRO = Rv!ID_REGISTRO
            Rs!FECHA_ENVIO = Rv!FECHA_ENVIO
            Rs!USUARIO_CREADOR = Rv!USUARIO_CREADOR
            Rs!ID_CLIENTE = Rv!ID_CLIENTE
            Rs!ID_ORDEN_BO = Rv!ID_ORDEN_BO
            Rs!TIPO_VENTA = Rv!TIPO_VENTA
            Rs!TIPO_ALTA = Rv!TIPO_ALTA
            Rs!PRODUCTO_CARACTERISTICA = Rv!PRODUCTO_CARACTERISTICA
            Rs!OFERTA_ASOCIADA = Rv!OFERTA_ASOCIADA
            Rs!OBSERVACIONES = Rv!OBSERVACIONES
            Rs!FECHA_GESTION = Rv!FECHA_GESTION
            Rs.Update
            
            Rv.Delete
            Rv.MoveNext

Rv.Close
Set Rv = Nothing

ListBox8.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_GESTION,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_BO,TIPO_VENTA FROM TB_Venta_Curso ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox8.AddItem Rs!ID_REGISTRO
Me.ListBox8.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox8.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox8.List(a, 3) = Rs!ID_ORDEN_BO
Me.ListBox8.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox8.List(a, 5) = Rs!FECHA_GESTION
a = a + 1
Rs.MoveNext
Loop

ListBox9.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox9.AddItem Rs!ID_REGISTRO
Me.ListBox9.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox9.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox9.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox9.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox9.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS2 FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label489 = Rs("CONTAR_REGISTROS2")

Rs.Close
Set Rs = Nothing
miConexion.Close

CommandButton75.Enabled = False
CommandButton75.BackColor = &HE0E0E0

Label481 = ListBox8.ListCount
Label491 = ListBox9.ListCount

Label502 = "Venta obtenida correctamente en su inbox"
Label502.ForeColor = &H0&

End If

Exit Sub

End Sub



Private Sub CommandButton76_Click()

ListBox6.Clear

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN,RESULTADO,RESULTADO_FINAL,FECHA_FIN_GESTION,USUARIO_GESTOR FROM TB_Gestionados_EX WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
While Rg.EOF = False

Me.ListBox6.AddItem Rg!ID_REGISTRO
Me.ListBox6.List(Z, 1) = Rg!USUARIO_CREADOR
Me.ListBox6.List(Z, 2) = Rg!ID_CLIENTE
Me.ListBox6.List(Z, 3) = Rg!ID_ORDEN
Me.ListBox6.List(Z, 4) = Rg!RESULTADO
Me.ListBox6.List(Z, 5) = Rg!RESULTADO_FINAL
Me.ListBox6.List(Z, 6) = Rg!FECHA_FIN_GESTION
Z = Z + 1
Rg.MoveNext
Wend

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionados_EX WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label201 = Rg("CONTAR_REGISTROS")

Call BD_Gestionados
Set Rg = New ADODB.Recordset
Rg.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionados_FL WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox96.Text & "# AND #" & Label428.Caption & "#", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label191 = Rg("CONTAR_REGISTROS")

Label205.Caption = Val(Label201.Caption) + Val(Label191.Caption)

Rg.Close
Set Rg = Nothing
miConexion.Close

CommandButton40.Enabled = True
CommandButton40.BackColor = &HC0&
CommandButton41.Enabled = False
CommandButton41.BackColor = &HE0E0E0

Label199 = "Mis ventas gestionadas como exitosas"
Label199.ForeColor = &H0&

UserForm1.Caption = "TP Ventas vodafone - Mis gestiones"
MultiPage1.Value = 4

End Sub

Private Sub CommandButton9_Click()

'On Error GoTo Error_obtener

If Label17 = "5" Then
Label21.ForeColor = &HC0&
Label21 = "Superó el límite de ventas permitidas que puede obtener"
Exit Sub

Else

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM TB_Principal", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT * FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText


            Rs.MoveFirst
            Rs.Find " ID_REGISTRO = '" & ListBox1.Text & "'"
            
            Rv.AddNew
            Rv!ID_REGISTRO = Rs!ID_REGISTRO
            Rv!FECHA_ENVIO = Rs!FECHA_ENVIO
            Rv!USUARIO_CREADOR = Rs!USUARIO_CREADOR
            Rv!ID_CLIENTE = Rs!ID_CLIENTE
            Rv!ID_ORDEN_INTERACCION = Rs!ID_ORDEN_INTERACCION
            Rv!TIPO_VENTA = Rs!TIPO_VENTA
            Rv!TIPO_ALTA = Rs!TIPO_ALTA
            Rv!PRODUCTO_CARACTERISTICA = Rs!PRODUCTO_CARACTERISTICA
            Rv!OFERTA_ASOCIADA = Rs!OFERTA_ASOCIADA
            Rv!OBSERVACIONES = Rs!OBSERVACIONES
            Rv!USUARIO_GESTOR = Label22.Caption
            Rv.Update
            
            Rs.Delete
            Rs.MoveNext

Rv.Close
Set Rv = Nothing

ListBox1.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Principal ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox1.AddItem Rs!ID_REGISTRO
Me.ListBox1.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox1.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox1.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox1.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox1.List(a, 5) = Rs!FECHA_ENVIO
a = a + 1
Rs.MoveNext
Loop

ListBox2.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox2.AddItem Rs!ID_REGISTRO
Me.ListBox2.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox2.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox2.List(c, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox2.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox2.List(c, 5) = Rs!FECHA_ENVIO
c = c + 1
Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
miConexion.Close

CommandButton9.Enabled = False
CommandButton9.BackColor = &HE0E0E0

Label17 = ListBox2.ListCount
Label14 = ListBox1.ListCount

Label21 = "Venta obtenida correctamente en su inbox"
Label21.ForeColor = &H0&
End If

Exit Sub
            
'___________________________________ Error al obtener un registro ________________________________________________________________
            
Error_obtener:
On Error GoTo Error_obtener2

                
Label21 = "La venta fue obtenida por otro usuario. Por favor, inténtalo de nuevo"
Label21.ForeColor = &HC0&

ListBox1.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM TB_Principal ORDER BY FECHA_ENVIOASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Do Until Rs.EOF
Me.ListBox1.AddItem Rs!FECHA_Y_HORA
Me.ListBox1.List(d, 1) = Rs!USUARIO_AGENTE
Me.ListBox1.List(d, 2) = Rs!NOMBRE_Y_APELLIDO
Me.ListBox1.List(d, 3) = Rs!NUMERO_DOCUMENTO
Me.ListBox1.List(d, 4) = Rs!MOVIL_CONTACTO
Me.ListBox1.List(d, 5) = Rs!PRODUCTO_VENTA
d = d + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing

Call BD_Principal
Set Rv = New ADODB.Recordset
Rv.Open "SELECT * FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rv.Close
Set Rv = Nothing
miConexion.Close

Label14 = ListBox1.ListCount

CommandButton9.Enabled = False
CommandButton9.BackColor = &HE0E0E0

Exit Sub
Error_obtener2:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Image11_Click()

''On Error GoTo error_Handler

Label341 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))

ListBox1.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Principal ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox1.AddItem Rs!ID_REGISTRO
Me.ListBox1.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox1.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox1.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox1.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox1.List(a, 5) = Rs!FECHA_ENVIO
a = a + 1
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing

ListBox2.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox2.AddItem Rs!ID_REGISTRO
Me.ListBox2.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox2.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox2.List(c, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox2.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox2.List(c, 5) = Rs!FECHA_ENVIO
c = c + 1
Rs.MoveNext
Wend

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS2 FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label125 = Rs("CONTAR_REGISTROS2")

Rs.Close
Set Rs = Nothing
miConexion.Close

'------------------------------------------------------------------

Label14 = ListBox1.ListCount
Label17 = ListBox2.ListCount

CommandButton9.Enabled = False
CommandButton14.Enabled = False

CommandButton9.BackColor = &HE0E0E0
CommandButton14.BackColor = &HE0E0E0

CommandButton17.Enabled = False
CommandButton17.BackColor = &HE0E0E0
CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0

CommandButton23.Enabled = False
CommandButton23.BackColor = &HE0E0E0

Label21 = "TP ventas vodafone, Bandeja de entrada para nuevas ventas"
Label21.ForeColor = &H0&
MultiPage1.Value = 0
UserForm1.Caption = "TP Ventas vodafone - Bandeja de entrada"
Label341 = "Bandeja de entrada"

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Image12_Click()

'On Error GoTo error_Handler

Label342 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))

CommandButton33.Enabled = False
CommandButton33.BackColor = &HE0E0E0
CommandButton37.Enabled = False
CommandButton37.BackColor = &HE0E0E0

Label136 = "Ventas agendadas pendientes por gestionar"
Label136.ForeColor = &H0&

'_______________________________ Ventas reagendadas __________________________________________

ListBox4.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Agendado ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox4.AddItem Rs!ID_REGISTRO
Me.ListBox4.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox4.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox4.List(a, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox4.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox4.List(a, 5) = Rs!FECHA_AGENDADO
a = a + 1
Rs.MoveNext
Loop

ListBox5.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA,FECHA_AGENDADO FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_AGENDADO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox5.AddItem Rs!ID_REGISTRO
Me.ListBox5.List(i, 1) = Rs!USUARIO_CREADOR
Me.ListBox5.List(i, 2) = Rs!ID_CLIENTE
Me.ListBox5.List(i, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox5.List(i, 4) = Rs!TIPO_VENTA
Me.ListBox5.List(i, 5) = Rs!FECHA_AGENDADO
i = i + 1
Rs.MoveNext
Loop


Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS2 FROM TB_Principal", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label138 = Rs("CONTAR_REGISTROS2")

Label128 = ListBox4.ListCount
Label142 = ListBox5.ListCount

Rs.Close
Set Rs = Nothing
miConexion.Close

MultiPage1.Value = 2
UserForm1.Caption = "TP Ventas vodafone - Ventas agendadas"
Label342 = "Ventas agendadas"

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Image13_Click()

'On Error GoTo error_Handler

Label347 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))

Label80 = "Mi inbox, ventas pendientes por gestionar"
Label80.ForeColor = &H0&

Label182 = "Mi inbox, ventas pendientes por gestionar"

Label76.BorderColor = &H80000001
Label95.BorderColor = &H808080
Label478.BorderColor = &H808080

ListBox3.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_ENVIO,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_INTERACCION,TIPO_VENTA FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_ENVIO ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox3.AddItem Rs!ID_REGISTRO
Me.ListBox3.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox3.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox3.List(c, 3) = Rs!ID_ORDEN_INTERACCION
Me.ListBox3.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox3.List(c, 5) = Rs!FECHA_ENVIO
c = c + 1
Rs.MoveNext
Wend

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT COUNT(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label76 = Rs("CONTAR_REGISTROS")

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) AS CONTAR_REGISTROS FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label95 = Rs("CONTAR_REGISTROS")

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Label478 = Rs("CONTAR_REGISTROS")

Rs.Close
Set Rs = Nothing
miConexion.Close

Label89 = "ID de registro"
Label88 = "Usuario creador"
Label87 = "ID del cliente"
Label86 = "ID de orden"
Label85 = "Tipo de venta"
Label84 = "Fecha de envío"

Label49 = "ID de orden"

CommandButton17.Enabled = False
CommandButton17.BackColor = &HE0E0E0
CommandButton23.Enabled = True
CommandButton23.BackColor = &HC0&
CommandButton22.Enabled = True
CommandButton22.BackColor = &HC0&


'__________________________________________________________________________

CommandButton18.Enabled = False
CommandButton18.BackColor = &HE0E0E0
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0

ComboBox1.Enabled = False
ComboBox1.BackColor = &HE0E0E0
ComboBox2.Enabled = False
ComboBox2.BackColor = &HE0E0E0
ComboBox3.Enabled = False
ComboBox3.BackColor = &HE0E0E0

TextBox5.Enabled = False
TextBox5.BackColor = &HE0E0E0
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
TextBox6.Enabled = False
TextBox6.BackColor = &HE0E0E0

TextBox01 = ""
TextBox02 = ""
TextBox03 = ""
TextBox04 = ""
TextBox05 = ""
TextBox06 = ""
TextBox07 = ""
TextBox08 = ""
TextBox09 = ""
Label58 = ""

TextBox01.Enabled = False
TextBox02.Enabled = False
TextBox03.Enabled = False
TextBox04.Enabled = False
TextBox05.Enabled = False
TextBox06.Enabled = False
TextBox07.Enabled = False
TextBox08.Enabled = False
TextBox09.Enabled = False
Label58.Enabled = False


MultiPage1.Value = 1
MultiPage2.Value = 0
MultiPage3.Value = 0

Label347 = "Ir a mi inbox"
UserForm1.Caption = "TP Ventas vodafone - Mi inbox"

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Image14_Click()

Label368 = "Cargando..."
TextBox96 = ""
TextBox97 = ""
Application.Wait (Now + TimeValue("00:00:01"))
MultiPage1.Value = 11
Label368 = "Ventas finalizadas"

End Sub

Private Sub Image15_Click()

Application.Wait (Now + TimeValue("00:00:02"))

MultiPage1.Value = 7
UserForm1.Caption = "TP Ventas vodafone - Cambiar contraseña"
TextBox87.Text = ""
TextBox85.Text = ""
TextBox86.Text = ""
TextBox87.SetFocus

End Sub

Private Sub Image34_Click()

'On Error GoTo error_Handler

Label382 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))

MultiPage5.Value = 0
TextBox88.Text = Date
TextBox88 = Format(TextBox88, "mm/dd/yyyy")
TextBox89.Text = Date
TextBox89 = Format(TextBox89, "mm/dd/yyyy")
Label428.Caption = Date + 1
Label428 = Format(Label428, "mm/dd/yyyy")

ListBox7.Clear

Call BD_Productividad
Set Pr = New ADODB.Recordset
Pr.Open "SELECT ID_REGISTRO, USUARIO_GESTOR, ID_CLIENTE, RESULTADO, RESULTADO_FINAL, ID_ORDEN_BO, TIEMPO_GESTION, FECHA_FIN_GESTION FROM TB_Productividad WHERE USUARIO_GESTOR = '" & Label22.Caption & "' AND FECHA_FIN_GESTION BETWEEN #" & TextBox88.Text & "# AND #" & Label428.Caption & "# ORDER BY FECHA_FIN_GESTION DESC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Pr.EOF = False
Me.ListBox7.AddItem Pr!ID_REGISTRO
Me.ListBox7.List(e, 1) = Pr!FECHA_FIN_GESTION
Me.ListBox7.List(e, 2) = Pr!ID_CLIENTE
Me.ListBox7.List(e, 3) = Pr!RESULTADO
Me.ListBox7.List(e, 4) = Pr!RESULTADO_FINAL
Me.ListBox7.List(e, 5) = Pr!TIEMPO_GESTION
e = e + 1
Pr.MoveNext
Wend

Pr.Close
Set Pr = Nothing
miConexion.Close

Label527 = ListBox7.ListCount

If Label527 <> "0" Then
CommandButton70.Enabled = True
CommandButton70.BackColor = &HC0&
Else
CommandButton70.Enabled = False
CommandButton70.BackColor = &HE0E0E0
End If

Label382 = "Productividad"
MultiPage1.Value = 8

Label392 = "Historial de ventas gestionadas el dia de hoy " & Date & ""
Label392.ForeColor = &H0&

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub

Private Sub Image35_Click()

Label425 = "Fecha inicio"
Load frmCalendario
frmCalendario.Show

 
End Sub

Private Sub Image36_Click()

Label425 = "Fecha fin"
Load frmCalendario
frmCalendario.Show

End Sub

Private Sub Image37_Click()

txtPassword = ""
MultiPage1.Value = 5
txtPassword.SetFocus

End Sub

Private Sub Image41_Click()
MultiPage1.Value = 1
MultiPage2.Value = 0
MultiPage3.Value = 0
End Sub

Private Sub Image42_Click()

Label508.Visible = True
Application.Wait (Now + TimeValue("00:00:01"))

TextBox001.Text = TextBox01.Text
TextBox002.Text = TextBox02.Text
TextBox003.Text = TextBox03.Text
ComboBox004.Text = TextBox04.Text
ComboBox005.Text = TextBox05.Text
ComboBox006.Text = TextBox06.Text
ComboBox007.Text = TextBox07.Text
TextBox008.Text = TextBox08.Text
TextBox009.Text = TextBox09.Text

Label456.Caption = Label58.Caption


Label508.Visible = False
MultiPage2.Value = 2
ComboBox004.SetFocus

End Sub

Private Sub Image45_Click()

If ComboBox004.Text = "" Or ComboBox005.Text = "" Or ComboBox006.Text = "" Or ComboBox007.Text = "" Then

Label462 = "Error al guardar"
Label462.Visible = True

Else

Label462 = "Guardando..."
Label462.Visible = True
Application.Wait (Now + TimeValue("00:00:02"))

TextBox01.Text = TextBox001.Text
TextBox02.Text = TextBox002.Text
TextBox03.Text = TextBox003.Text
TextBox04.Text = ComboBox004.Text
TextBox05.Text = ComboBox005.Text
TextBox06.Text = ComboBox006.Text
TextBox07.Text = ComboBox007.Text
TextBox08.Text = TextBox008.Text
TextBox09.Text = TextBox009.Text
Label456.Caption = Label58.Caption

Label462.Visible = False
MultiPage2.Value = 0
End If

End Sub

Private Sub Image46_Click()

Label508.Visible = True
Application.Wait (Now + TimeValue("00:00:01"))

TextBox94.Text = ""
ComboBox8 = ""
TextBox93.Text = TextBox08.Text
Label508.Visible = False
MultiPage2.Value = 3
ComboBox8.SetFocus

End Sub

Private Sub Image47_Click()

Label462 = "Cancelando..."
Label462.Visible = True
Application.Wait (Now + TimeValue("00:00:02"))
Label462.Visible = False
MultiPage2.Value = 0

End Sub

Private Sub Image48_Click()

If TextBox94.Text = "" Or ComboBox8.Text = "" Then
Label474 = "Error al enviar"
Label474.Visible = True
Else

Label474 = "Enviando..."
Label474.Visible = True
Application.Wait (Now + TimeValue("00:00:02"))

    
    Call BD_Gestionados
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * FROM Usuaros_Reportados", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
    Rs.AddNew
    Rs.Fields("USUARIO_CREADOR") = Label78.Caption
    Rs.Fields("MOTIVO_REPORTE") = ComboBox8.Text
    Rs.Fields("USUARIO_REPORTADO") = TextBox08.Text
    Rs.Fields("OBSERVACIONES") = TextBox94.Text
    Rs.Fields("ID_REGISTRO") = TextBox09.Text
    Rs.Update
    Label474.Visible = False
    MultiPage2.Value = 0
    
Rs.Close
Set Rs = Nothing
miConexion.Close
     
End If

End Sub

Private Sub Image49_Click()
Label474 = "Cancelando..."
Label474.Visible = True
Application.Wait (Now + TimeValue("00:00:02"))
Label474.Visible = False
MultiPage2.Value = 0
End Sub

Private Sub Image50_Click()

'On Error GoTo error_Handler

Label506 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))

ListBox8.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO,FECHA_GESTION,USUARIO_CREADOR,ID_CLIENTE,ID_ORDEN_BO,TIPO_VENTA FROM TB_Venta_Curso ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
Me.ListBox8.AddItem Rs!ID_REGISTRO
Me.ListBox8.List(a, 1) = Rs!USUARIO_CREADOR
Me.ListBox8.List(a, 2) = Rs!ID_CLIENTE
Me.ListBox8.List(a, 3) = Rs!ID_ORDEN_BO
Me.ListBox8.List(a, 4) = Rs!TIPO_VENTA
Me.ListBox8.List(a, 5) = Rs!FECHA_GESTION
a = a + 1
Rs.MoveNext
Loop

ListBox9.Clear

Call BD_Principal
Set Rs = New ADODB.Recordset
Rs.Open "SELECT ID_REGISTRO, FECHA_GESTION, USUARIO_CREADOR, ID_CLIENTE, ID_ORDEN_BO, TIPO_VENTA, USUARIO_GESTOR FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "' ORDER BY FECHA_GESTION ASC", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

While Rs.EOF = False
Me.ListBox9.AddItem Rs!ID_REGISTRO
Me.ListBox9.List(c, 1) = Rs!USUARIO_CREADOR
Me.ListBox9.List(c, 2) = Rs!ID_CLIENTE
Me.ListBox9.List(c, 3) = Rs!ID_ORDEN_BO
Me.ListBox9.List(c, 4) = Rs!TIPO_VENTA
Me.ListBox9.List(c, 5) = Rs!FECHA_GESTION
c = c + 1
Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
miConexion.Close

'------------------------------------------------------------------

Label481 = ListBox8.ListCount
Label491 = ListBox9.ListCount

CommandButton72.Enabled = False
CommandButton75.Enabled = False

CommandButton72.BackColor = &HE0E0E0
CommandButton75.BackColor = &HE0E0E0

Label502 = "TP ventas vodafone, Bandeja de ventas en curso"
Label502.ForeColor = &H0&
MultiPage1.Value = 10
UserForm1.Caption = "TP Ventas vodafone - Ventas en curso"
Label506 = "Ventas en curso"

End Sub



Private Sub Image61_Click()

Label425 = "Fecha inicio X"
Load frmCalendario
frmCalendario.Show

End Sub

Private Sub Image62_Click()

Label425 = "Fecha fin X"
Load frmCalendario
frmCalendario.Show

End Sub

Private Sub Image65_Click()
MultiPage1.Value = 6
End Sub

Private Sub Label217_Click()
MultiPage4.Value = 1
TextBox81 = Environ("USERNAME")
TextBox81.Enabled = False
End Sub

Private Sub Label221_Click()
MultiPage4.Value = 0
  TextBox80 = ""
  TextBox81 = ""
  TextBox82 = ""
  TextBox83 = ""
  ComboBox4 = ""
End Sub



Private Sub Label341_Click()
Image11_Click
End Sub

Private Sub Label342_Click()
Image12_Click
End Sub

Private Sub Label347_Click()
Image13_Click
End Sub



Private Sub Label356_Click()
 UserForm1.Caption = "TP Ventas vodafone - Menú principal"
 MultiPage1.Value = 6
End Sub

Private Sub Label368_Click()
Image14_Click
End Sub

Private Sub Label372_Click()
Image15_Click
End Sub

Private Sub Label378_Click()

'On Error GoTo error_Handler

Dim Gestionando(1 To 3) As Integer

    Label378 = "Cargando..."
    Application.Wait (Now + TimeValue("00:00:01"))

    UserForm1.Caption = "TP Ventas vodafone - Menú principal"
    MultiPage1.Value = 6
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS FROM TB_Gestionando WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Gestionando(1) = Ra("CONTAR_REGISTROS")
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS1 FROM TB_Gestionando_Agendado WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Gestionando(2) = Ra("CONTAR_REGISTROS1")
    
    Call BD_Principal
    Set Ra = New ADODB.Recordset
    Ra.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS1 FROM TB_Gestionando_Venta_Curso WHERE USUARIO_GESTOR = '" & Label22.Caption & "'", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    Gestionando(3) = Ra("CONTAR_REGISTROS1")
    
    
    Label371.Caption = Val(Gestionando(1)) + Val(Gestionando(2)) + Val(Gestionando(3))
    
    
    
    Ra.Close
    Set Ra = Nothing
  
     Call BD_Principal
     Set Re = New ADODB.Recordset
     Re.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS2 FROM TB_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
     Label370 = Re("CONTAR_REGISTROS2")
     
     Re.Close
     Set Re = Nothing
        
     Call BD_Principal
     Set Rs = New ADODB.Recordset
     Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS3 FROM TB_Principal", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
     Label369 = Rs("CONTAR_REGISTROS3")
     
     Call BD_Principal
     Set Rs = New ADODB.Recordset
     Rs.Open "SELECT Count(ID_REGISTRO) as CONTAR_REGISTROS4 FROM TB_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
     Label507 = Rs("CONTAR_REGISTROS4")
     
   
     
     Rs.Close
     Set Rs = Nothing
     miConexion.Close
     
     Label378 = "Ir a inicio"
     Label379 = "Ir a inicio"
     Label380 = "Ir a inicio"
     Label381 = "Ir a inicio"
     Label401 = "Ir a inicio"
     Label504 = "Ir a inicio"
     
     
Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close
     
     
End Sub

Private Sub Label379_Click()
Label379 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))
Label378_Click
End Sub

Private Sub Label380_Click()
Label380 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))
Label378_Click
End Sub

Private Sub Label381_Click()
Label381 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))
Label378_Click
End Sub

Private Sub Label382_Click()
Image34_Click
End Sub

Private Sub Label401_Click()
Label401 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))
Label378_Click
End Sub

Private Sub Label429_Click()

End Sub

Private Sub Label504_Click()

Label504 = "Cargando..."
Application.Wait (Now + TimeValue("00:00:01"))
Label378_Click

End Sub

Private Sub Label58_Click()
MultiPage2.Value = 1
TextBox77.Text = Label58.Caption
End Sub



Private Sub ListBox1_Click()

CommandButton9.Enabled = True
CommandButton9.BackColor = &HC0&
CommandButton14.Enabled = False
CommandButton14.BackColor = &HE0E0E0
Label21 = "Venta seleccionada en la bandeja de entrada"
Label21.ForeColor = &H0&

End Sub

Private Sub ListBox2_Click()

CommandButton9.Enabled = False
CommandButton9.BackColor = &HE0E0E0

CommandButton14.Enabled = True
CommandButton14.BackColor = &HC0&

Label21 = "Venta seleccionada en su inbox"
Label21.ForeColor = &H0&

End Sub

Private Sub ListBox3_Click()

On Error GoTo error_Handler

If Label182 = "Mi inbox, ventas pendientes por gestionar" Then

Call BD_Principal
Set Rx = New ADODB.Recordset
Rx.Open "SELECT * FROM TB_Gestionando", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rx.MoveFirst
Rx.Find " ID_REGISTRO = '" & ListBox3.Text & "'"
TextBox01.Text = Rx!FECHA_ENVIO
TextBox02.Text = Rx!ID_CLIENTE
TextBox03.Text = Rx!ID_ORDEN_INTERACCION
TextBox04.Text = Rx!TIPO_VENTA
TextBox05.Text = Rx!TIPO_ALTA
TextBox06.Text = Rx!PRODUCTO_CARACTERISTICA
TextBox07.Text = Rx!OFERTA_ASOCIADA
TextBox08.Text = Rx!USUARIO_CREADOR
TextBox09.Text = Rx!ID_REGISTRO
Label58.Caption = Rx!OBSERVACIONES

Rx.Close
Set Rx = Nothing
miConexion.Close

End If

If Label182 = "Mi inbox, ventas agendadas pendientes por gestionar" Then

Call BD_Principal
Set Rx = New ADODB.Recordset
Rx.Open "SELECT * FROM TB_Gestionando_Agendado", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rx.MoveFirst
Rx.Find " ID_REGISTRO = '" & ListBox3.Text & "'"
TextBox01.Text = Rx!FECHA_ENVIO
TextBox02.Text = Rx!ID_CLIENTE
TextBox03.Text = Rx!ID_ORDEN_INTERACCION
TextBox04.Text = Rx!TIPO_VENTA
TextBox05.Text = Rx!TIPO_ALTA
TextBox06.Text = Rx!PRODUCTO_CARACTERISTICA
TextBox07.Text = Rx!OFERTA_ASOCIADA
TextBox08.Text = Rx!USUARIO_CREADOR
TextBox09.Text = Rx!ID_REGISTRO
Label58.Caption = Rx!OBSERVACIONES

Rx.Close
Set Rx = Nothing
miConexion.Close

End If

If Label182 = "Mi inbox, ventas en curso pendientes por gestionar" Then

Call BD_Principal
Set Rx = New ADODB.Recordset
Rx.Open "SELECT * FROM TB_Gestionando_Venta_Curso", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Rx.MoveFirst
Rx.Find " ID_REGISTRO = '" & ListBox3.Text & "'"
TextBox01.Text = Rx!FECHA_ENVIO
TextBox02.Text = Rx!ID_CLIENTE
TextBox03.Text = Rx!ID_ORDEN_BO
TextBox04.Text = Rx!TIPO_VENTA
TextBox05.Text = Rx!TIPO_ALTA
TextBox06.Text = Rx!PRODUCTO_CARACTERISTICA
TextBox07.Text = Rx!OFERTA_ASOCIADA
TextBox08.Text = Rx!USUARIO_CREADOR
TextBox09.Text = Rx!ID_REGISTRO
Label58.Caption = Rx!OBSERVACIONES

Rx.Close
Set Rx = Nothing
miConexion.Close

End If

MultiPage2.Value = 0

Exit Sub
error_Handler:
MultiPage1.Value = 3
Label180 = " " & Err.Description
Label180.ForeColor = &H0&
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"

Call BD_Seguridad
Set Rk = New ADODB.Recordset
Rk.Open "SELECT * FROM Recopilador_errores", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
Rk.AddNew
Rk.Fields("USUARIO") = Label22.Caption
Rk.Fields("DESCRIPCION") = Label180.Caption
Rk.Fields("APLICACION") = Label153.Caption
Rk.Update

Rk.Close
Set Rk = Nothing
miConexion.Close

End Sub


Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Inicio_Gestión = Time

If Label182 = "Mi inbox, ventas pendientes por gestionar" Then
Label80 = "Venta nueva, Gestión en curso..."
End If

If Label182 = "Mi inbox, ventas agendadas pendientes por gestionar" Then
Label80 = "Venta agendada, Gestión en curso..."
End If

If Label182 = "Mi inbox, ventas en curso pendientes por gestionar" Then
Label80 = "Venta en curso, Gestión en curso..."
End If

CommandButton17.Enabled = False
CommandButton17.BackColor = &HE0E0E0
CommandButton22.Enabled = False
CommandButton22.BackColor = &HE0E0E0
CommandButton23.Enabled = False
CommandButton23.BackColor = &HE0E0E0
CommandButton21.Enabled = False
CommandButton21.BackColor = &HE0E0E0

ListBox3.Enabled = False
ListBox3.BackColor = &HE0E0E0

ComboBox1.Enabled = True
ComboBox1.BackColor = &H80000018

TextBox01.Enabled = True
TextBox02.Enabled = True
TextBox03.Enabled = True
TextBox04.Enabled = True
TextBox05.Enabled = True
TextBox06.Enabled = True
TextBox07.Enabled = True
TextBox08.Enabled = True
TextBox09.Enabled = True
Label58.Enabled = True

Label379.Visible = False

CommandButton18.Enabled = True
CommandButton18.BackColor = &HC0&

Image42.Visible = True
Image46.Visible = True


End Sub

Private Sub ListBox4_Click()

CommandButton33.Enabled = True
CommandButton33.BackColor = &HC0&

CommandButton37.Enabled = False
CommandButton37.BackColor = &HE0E0E0

Label136 = "Venta seleccionada, haga clic en obtener venta agendada"
Label136.ForeColor = &H0&

End Sub

Private Sub ListBox5_Click()


CommandButton37.Enabled = True
CommandButton37.BackColor = &HC0&

CommandButton33.Enabled = False
CommandButton33.BackColor = &HE0E0E0

Label136 = "Venta seleccionada en su inbox"
Label136.ForeColor = &H0&

End Sub

Private Sub ListBox6_Click()

End Sub

Private Sub ListBox8_Click()

CommandButton72.Enabled = True
CommandButton72.BackColor = &HC0&
CommandButton75.Enabled = False
CommandButton75.BackColor = &HE0E0E0
Label502 = "Venta seleccionada en la bandeja de ventas en curso"
Label502.ForeColor = &H0&

End Sub

Private Sub ListBox9_Click()

CommandButton75.Enabled = True
CommandButton75.BackColor = &HC0&
CommandButton72.Enabled = False
CommandButton72.BackColor = &HE0E0E0
Label502 = "Venta seleccionada en su inbox"
Label502.ForeColor = &H0&

End Sub


Private Sub MultiPage5_Change()

End Sub

Private Sub TextBox2_Change()

TextBox2 = UCase(Left(TextBox2, 1)) & LCase(Mid(TextBox2, 2, Len(TextBox2)))

If TextBox2 = "" Then
CommandButton19.Enabled = False
CommandButton19.BackColor = &HE0E0E0
Else
CommandButton19.Enabled = True
CommandButton19.BackColor = &HC0&
End If

End Sub

Private Sub TextBox5_Change()

TextBox5.Text = Trim$(QuitaEspacios(TextBox5.Text, True, True))

If ComboBox1 = "Venta en curso" Then
If TextBox5.Text = "" Then
TextBox95 = ""
TextBox95.Enabled = False
TextBox95.BackColor = &HE0E0E0
Else
TextBox95.Enabled = True
TextBox95.BackColor = &H80000018
TextBox95 = ""
TextBox95 = Now
End If
End If

If ComboBox1 <> "Venta en curso" Then
If TextBox5.Text = "" Then
TextBox2 = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
Else
TextBox2.Enabled = True
TextBox2.BackColor = &H80000018
End If
End If

End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
 KeyAscii = KeyAscii
 Else
 KeyAscii = 0
 End If
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox6.BackColor = &H80000018

If TextBox6.Text = "" Then
TextBox2 = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
Else
TextBox2.Enabled = True
TextBox2.BackColor = &H80000018
End If

End Sub

Private Sub TextBox80_Change()
TextBox80 = UCase(TextBox80)
End Sub

Private Sub TextBox81_Change()
 TextBox81 = UCase(TextBox81)
End Sub

Private Sub TextBox85_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
         CommandButton64.SetFocus
        CommandButton64_Click
        End If
End Sub

Private Sub TextBox86_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
         CommandButton64.SetFocus
        CommandButton64_Click
        End If
End Sub

Private Sub TextBox87_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
         CommandButton64.SetFocus
        CommandButton64_Click
        End If
End Sub

Private Sub TextBox94_Change()
Label474.Visible = False
TextBox94 = UCase(Left(TextBox94, 1)) & LCase(Mid(TextBox94, 2, Len(TextBox94)))
End Sub

Private Sub TextBox95_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If TextBox95.Text = "" Then
TextBox2 = ""
TextBox2.Enabled = False
TextBox2.BackColor = &HE0E0E0
Else
TextBox2.Enabled = True
TextBox2.BackColor = &H80000018
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
UserForm1.Hide

End

End If

Cancel = True

End Sub


Private Sub UserForm_Initialize()

On Error GoTo error_Handler:

With ComboBox1
.AddItem "Venta en curso"
.AddItem "Agendar venta"
.AddItem "Venta exitosa"
.AddItem "Venta fallida"
End With

With ComboBox3
.AddItem "Enviar a mi inbox"
.AddItem "Enviar a la bandeja"
End With

With ComboBox4
 .AddItem "Vodafone datos"
 .AddItem "Vodafone Fibra"
 .AddItem "Vodafone Adsl"
 .AddItem "Vodafone 123"
 .AddItem "Vodafone Bajas"
End With

With ComboBox8
.AddItem "Mejoras de ofertas"
.AddItem "Renovación descuento"
.AddItem "Traslado de domicilio"
.AddItem "Venta sin cargar en Smart"
.AddItem "Cliente no reconoce oferta"
.AddItem "Error de información/Incompleta"
End With
  
txtUsuario = Environ("USERNAME")
txtUsuario.Enabled = False
ActualizarComboTipoVenta

UserForm1.Caption = "TP Ventas vodafone - Inicio de sesión"
MultiPage1.Value = 5
MultiPage4.Value = 0


Exit Sub
error_Handler:
Error_Sistema = Err.Description
MultiPage1.Value = 3
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"
Application.Wait (Now + TimeValue("00:00:03"))

Error_sistemas

End Sub


Sub ActualizarComboTipoVenta()

On Error GoTo error_Handler:

ComboBox004.Clear

Call BD_Tipificacion
Set Rs = New ADODB.Recordset
Rs.Open "SELECT DISTINCT(TIPO_VENTA) FROM Productos_items", miConexion, adOpenKeyset, adLockOptimistic, adCmdText

Do Until Rs.EOF
ComboBox004.AddItem Rs.Fields(0)
Rs.MoveNext
Loop

Rs.Close
Set Rs = Nothing
miConexion.Close

Exit Sub
error_Handler:
Error_Sistema = Err.Description
MultiPage1.Value = 3
UserForm1.Caption = "TP Ventas vodafone - Error inesperado"
Application.Wait (Now + TimeValue("00:00:03"))

Error_sistemas

End Sub


Function QuitaEspacios(ByVal Texto As String, ByVal DelComienzo As Boolean, ByVal DelFinal As Boolean) As String

On Local Error Resume Next

If DelComienzo = True Then
Do Until InStr(1, Texto, vbCrLf & " ") = 0
Texto = Replace(Texto, vbCrLf & " ", vbCrLf)
Loop
End If

If DelFinal = True Then
Do Until InStr(1, Texto, " " & vbCrLf) = 0
Texto = Replace(Texto, " " & vbCrLf, vbCrLf)
Loop
End If

QuitaEspacios = Texto

End Function

