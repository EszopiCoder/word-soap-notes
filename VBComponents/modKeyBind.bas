Attribute VB_Name = "modKeyBind"
Option Explicit
 
Private Sub AddKeyBinding()
    With Application
        ' Do customization in ThisDocument
        .CustomizationContext = ThisDocument
         
        ' Add keybinding to this document Shortcut: Alt+0
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey0), _
            KeyCategory:=wdKeyCategoryCommand, _
            command:="TestKeybinding"
    End With
End Sub
Public Sub TestKeybinding()
    MsgBox "We have a winner", vbInformation, "Success"
End Sub

Public Sub showKeyBindForm()
    If Application.KeyBindings.Count = 0 Then
        MsgBox "No key bindings exist.", vbExclamation
    Else
        formKeyBind.Show
    End If
End Sub
Public Sub TemplateKeyBinding()
    With Application
         ' Do customization in ThisDocument
        .CustomizationContext = ThisDocument
         
        .KeyBindings.Add KeyCode:=wdKeyReturn, _
            KeyCategory:=wdKeyCategoryCommand, _
            command:="InsertTemplate"
    End With
End Sub
Public Sub RemoveKeyBinding()
    Dim i As Long
    
    With Application
        .CustomizationContext = ThisDocument
        For i = 1 To .KeyBindings.Count
            If .KeyBindings.Item(i).KeyCode = wdKeyReturn Then
                MsgBox "Key binding is deactivated for '" & _
                .KeyBindings.Item(i).KeyString & "'.", vbInformation
                .KeyBindings.Item(i).Clear
                Exit For
            End If
        Next i
    End With
End Sub
Public Sub InsertTemplate()
    Dim TemplateName As String
    Dim TemplateText As String
    Dim currentPosition As Range
    Set currentPosition = Selection.Range
    
    ' Note: Shortcut names must begin with a token character (such as #)
    
    Application.ScreenUpdating = False
    
    If ActiveDocument.Range.Characters.Count = 1 Then
        ' Perform function of return key if document is empty
        GoTo DefaultFx
    Else
        ' Return shortcut name                                      ' Comments:
        Selection.MoveLeft wdCharacter, _
            currentPosition.Information(wdFirstCharacterColumnNumber), _
            wdExtend                                                ' Select line before cursor
        TemplateName = Selection.Range.Text                         ' Set entire line to variable
        If InStrRev(TemplateName, "#") = 0 Then GoTo DefaultFx      ' Ensure token character is typed
        TemplateName = Right(TemplateName, _
            Len(TemplateName) - InStrRev(TemplateName, "#"))        ' Extract shortcut name
        currentPosition.Select                                      ' Return cursor to original position
        ' Find template in template dictionary
        ' If returns empty, no template exists
        TemplateText = GetTemplate(TemplateName)
        If Len(TemplateText) = 0 Then
            ' Perform function of return key
            GoTo DefaultFx
        End If
        ' Select and replace shortcut name with shortcut text       ' Comments:
        Selection.MoveLeft wdCharacter, _
            Len(TemplateName) + 1, wdExtend                         ' Select shortcut name
        Selection.Delete                                            ' Remove shortcut name
        Selection.TypeText TemplateText                             ' Insert shortcut text
    End If
    
    Set currentPosition = Nothing
    Application.ScreenUpdating = True
    Exit Sub

DefaultFx:
    ' Perform normal function and return screen updating to normal
    currentPosition.Select
    Selection.TypeText vbNewLine
    Set currentPosition = Nothing
    Application.ScreenUpdating = True
End Sub
