Attribute VB_Name = "modEvents"
Public myEvents As New clsEvents

Public Sub StartEvents()
    If myEvents.app Is Nothing Then
        Set myEvents.app = Word.Application
    End If
End Sub

Public Sub EndEvents()
    Set myEvents.app = Nothing
End Sub
