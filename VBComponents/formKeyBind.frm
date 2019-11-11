VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formKeyBind 
   Caption         =   "Manage Key Bindings"
   ClientHeight    =   2664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "formKeyBind.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formKeyBind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Add all key bindings to list
    Dim i As Long
    For i = 0 To Application.KeyBindings.Count - 1
        listKeyBind.AddItem
        listKeyBind.List(i, 0) = Application.KeyBindings.Item(i + 1).command
        listKeyBind.List(i, 1) = Application.KeyBindings.Item(i + 1).KeyString
    Next i
    ' Listbox settings
    With listKeyBind
        .MultiSelect = fmMultiSelectMulti
        .ColumnCount = 2
        .ColumnWidths = "130;60"
        .ListIndex = -1
    End With
End Sub

Private Sub btnDelete_Click()
    ' Declare variables
    Dim i As Long
    Dim selectCount As Long
    Dim retMsg As Long
    ' Determine number of key bindings
    For i = 0 To listKeyBind.ListCount - 1
        If listKeyBind.Selected(i) = True Then
            selectCount = selectCount + 1
        End If
    Next i
    ' Send message to user
    Select Case selectCount
        Case 0
            MsgBox "No item(s) selected.", vbExclamation
            Exit Sub
        Case 1
            retMsg = MsgBox("Are you sure you want to permanently delete " & _
                selectCount & " key binding?", vbYesNo + vbExclamation)
        Case Else
            retMsg = MsgBox("Are you sure you want to permanently delete these " & _
                selectCount & " key bindings?", vbYesNo + vbExclamation)
    End Select
    If retMsg = vbNo Then Exit Sub
    ' Delete key bindings
    For i = listKeyBind.ListCount - 1 To 0 Step -1
        If listKeyBind.Selected(i) = True Then
            listKeyBind.RemoveItem i
            Application.KeyBindings.Item(i + 1).Clear
        End If
    Next i
    MsgBox selectCount & " item(s) deleted.", vbInformation
    If listKeyBind.ListCount = 0 Then Unload Me
End Sub
