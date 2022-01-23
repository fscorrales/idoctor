Attribute VB_Name = "ConfigurarFormulario"
Function CenterMe(frmForm As Form)
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    frmForm.Top = (Screen.Height - frmForm.Height) / 2 - 500
End Function
