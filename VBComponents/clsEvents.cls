VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents app As Word.Application
Attribute app.VB_VarHelpID = -1

Private Sub app_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    If ThisDocument.Name <> ActiveDocument.Name Then
        ThisDocument.Saved = True
    End If
End Sub
