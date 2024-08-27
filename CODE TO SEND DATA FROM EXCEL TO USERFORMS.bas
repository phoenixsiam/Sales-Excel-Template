Attribute VB_Name = "Module7"
Sub SELL()
Attribute SELL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SELL Macro
'

'
    Range("Tabla5").Select
    Selection.ListObject.ListRows.Add (1)
    Range("G26:K26").Select
    Selection.Copy
    Range("G32").Select
    ActiveSheet.Paste
End Sub
