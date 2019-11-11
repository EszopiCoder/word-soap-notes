VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEditor 
   Caption         =   "Template Editor"
   ClientHeight    =   4092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9504
   OleObjectBlob   =   "formEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private listSkip As Boolean

Private Sub UserForm_Initialize()
    ' Check if dictionary is empty
    Dim varKey As Variant
    If objTemplates.Count > 0 Then
        ' Load dictionary to listbox
        For Each varKey In objTemplates.keys
            listTemplateName.AddItem varKey
        Next varKey
    Else
        MsgBox "No template dictionary found.", vbExclamation
    End If
    ' Set up userform
    textTemplate.ScrollBars = fmScrollBarsVertical
    textTemplate.WordWrap = True
    textTemplate.MultiLine = True
    textTemplate.EnterKeyBehavior = True
    Call LoadMode
End Sub
Private Sub listTemplateName_Click()
    ' Click event is trigger by listbox changes
    If listSkip = True Then Exit Sub
    
    Call EditMode
    
    ' Detect if any changes were not saved yet
    Call UnsavedTemplate(lblTemplateName.Caption, textTemplate.Text)
    
    ' Display template name
    lblTemplateName.Caption = listTemplateName.Value
    ' Display template in textbox
    textTemplate.Text = objTemplates(listTemplateName.Value)
End Sub
Private Sub btnDelete_Click()
    
    If listTemplateName.ListIndex = -1 Then
        MsgBox "No template selected.", vbExclamation
        Exit Sub
    End If
    listSkip = True
    Call RemoveTemplate(listTemplateName.Value)
    listTemplateName.RemoveItem listTemplateName.ListIndex
    listSkip = False
    Call LoadMode
End Sub
Private Sub btnSave_Click()
    ' Check if template is null
    If Len(textTemplate.Text) = 0 Then
        MsgBox "Template cannot be blank.", vbExclamation
        Exit Sub
    End If
    ' Add to dictionary and add to listbox if it doesn't exist
    If SaveTemplate(lblTemplateName.Caption, textTemplate.Text) = False Then
        listSkip = True
        listTemplateName.AddItem lblTemplateName.Caption
        listTemplateName.ListIndex = -1
        listSkip = False
    End If
    Call EditMode
End Sub
Private Sub btnAddTemplate_Click()
    
    Dim strName As String
    Call AddMode
    
    Do
        strName = InputBox("Enter template name.")
        If StrPtr(strName) = 0 Then
            Call LoadMode
            Exit Sub
        ElseIf strName = vbNullString Then
            MsgBox "Template name cannot be blank.", vbExclamation
        ElseIf objTemplates.Exists(strName) = True Then
            MsgBox "Template name """ & strName & """ already exists.", vbExclamation
        End If
    Loop Until objTemplates.Exists(strName) = False And strName <> vbNullString
    
    lblTemplateName.Caption = strName
End Sub
Private Sub btnCancel_Click()
    Call LoadMode
End Sub
Private Sub LoadMode()
    btnDelete.Enabled = False
    btnSave.Enabled = False
    btnAddTemplate.Enabled = True
    btnCancel.Enabled = False
    lblTemplateName.Caption = ""
    textTemplate.Text = ""
    textTemplate.Enabled = False
    listTemplateName.Enabled = True
    listSkip = True
    listTemplateName.ListIndex = -1
    listSkip = False
End Sub
Private Sub AddMode()
    btnDelete.Enabled = False
    btnSave.Enabled = True
    btnAddTemplate.Enabled = False
    btnCancel.Enabled = True
    lblTemplateName.Caption = ""
    textTemplate.Text = ""
    textTemplate.Enabled = True
    listTemplateName.Enabled = False
End Sub
Private Sub EditMode()
    btnDelete.Enabled = True
    btnSave.Enabled = True
    btnAddTemplate.Enabled = True
    btnCancel.Enabled = False
    textTemplate.Enabled = True
    listTemplateName.Enabled = True
End Sub

