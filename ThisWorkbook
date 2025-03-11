Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.DisplayFullScreen = False 'Ves pantalla completa
Application.DisplayFormulaBar = True 'Activa la barra de formulas
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'Cancel = True
End Sub

Private Sub Workbook_Open()
ActiveWindow.DisplayWorkbookTabs = False 'Oculta las fichas de las hojas
ActiveWindow.DisplayHeadings = False 'Oculta títulos
Application.DisplayFormulaBar = False 'Oculta la barra de formulas
ActiveWindow.DisplayGridlines = False 'Oculta las lineas de la cuadricula
Application.DisplayStatusBar = False 'Oculta la barra de estado
Application.DisplayFullScreen = True 'Ves pantalla completa
Application.Visible = True
End Sub

