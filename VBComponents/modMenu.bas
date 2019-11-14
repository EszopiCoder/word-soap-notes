Attribute VB_Name = "modMenu"
'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
'   <ribbon>
'      <tabs>
'         <tab id="TabSOAP" label="SOAP Notes">
'            <group id="SOAPTemplate" label="SOAP Templates">
'               <button id="KeyBindOn"
'                   imageMso = "ChangeToAcceptInvitation"
'                   label = "Activate"
'                   screentip="Activate Key Bind"
'                   supertip = "Turn key bind on."
'                   onAction = "KeyBindOn_Click"
'                   size="large"/>
'               <button id="KeyBindOff"
'                   imageMso = "ChangeToDeclineInvitation"
'                   label = "Deactivate"
'                   screentip="Deactivate Key Bind"
'                   supertip = "Turn key bind off."
'                   onAction = "KeyBindOff_Click"
'                   size="large"/>
'               <button id="LoadTemplate"
'                   imageMso = "ImportTextFile"
'                   label = "Load Template"
'                   screentip="Load Template"
'                   supertip = "Load template to dictionary."
'                   onAction = "LoadTemplate_Click"
'                   size="large"/>
'               <button id="SaveTemplate"
'                   imageMso = "FileSave"
'                   label = "Save Template"
'                   screentip="Save Template"
'                   supertip = "Save template to CSV."
'                   onAction = "SaveTemplate_Click"
'                   size="large"/>
'               <button id="TemplateEditor"
'                   imageMso = "DataFormWord"
'                   label = "Template Editor"
'                   screentip="Template Editor"
'                   supertip = "Open template editor."
'                   onAction = "TemplateEditor_Click"
'                   size="large"/>
'               <button id="getInfo"
'                   imageMso = "ARMPreviewButton"
'                   label = "Info"
'                   screentip="Information"
'                   supertip = "Return contact information."
'                   onAction = "getInfo_Click"
'                   size="large"/>
'            </group>
'         </tab>
'      </tabs>
'   </ribbon>
'</customUI>
'*********************************XML CODE*********************************

Option Explicit

Sub KeyBindOn_Click(control As IRibbonControl)
    Call modKeyBind.TemplateKeyBinding
    If modDictionary.DictionaryExists = False Then
        MsgBox "Key binding is activated." & vbNewLine & _
        "Template Detected: False", vbInformation
    Else
        MsgBox "Key binding is activated." & vbNewLine & _
        "Template Detected: True", vbInformation
    End If
    Call modEvents.StartEvents
End Sub
Sub KeyBindOff_Click(control As IRibbonControl)
    Call modKeyBind.RemoveKeyBinding
    Call modEvents.StartEvents
End Sub
Sub LoadTemplate_Click(control As IRibbonControl)
    Dim arrAddIn As AddIn
    Dim strPath As String
    Dim strFilter As String
    
    ' Retrieve strPath of add in
    For Each arrAddIn In AddIns
        If arrAddIn.Name = ThisDocument.Name Then
            strPath = arrAddIn.Path
        End If
    Next arrAddIn
    
    ' Open file dialog at the following path
    strFilter = modOpenDialog.OpenAddFilterItem(strFilter, "CSV (Comma delimited)", "*.csv")
    strPath = modOpenDialog.FileDialogOpen1(strPath, "Open CSV File", strFilter)
    
    ' Load template to dictionary
    Call modDictionary.OpenCSV(strPath)
    Call modEvents.StartEvents
End Sub
Sub SaveTemplate_Click(control As IRibbonControl)
    Dim arrAddIn As AddIn
    Dim strPath As String
    Dim strFilter As String
    
    ' Check if template dictionary exists
    If modDictionary.DictionaryExists = False Then
        MsgBox "No template dictionary loaded.", vbExclamation
        Exit Sub
    End If
    
    ' Check if template dictionary is empty
    If modDictionary.DictionaryCount = 0 Then
        MsgBox "Template dictionary is empty.", vbExclamation
        Exit Sub
    End If
    
    ' Retrieve strPath of add in
    For Each arrAddIn In AddIns
        If arrAddIn.Name = ThisDocument.Name Then
            strPath = arrAddIn.Path
        End If
    Next arrAddIn
    
    ' Save file dialog at the following path
    strFilter = modSaveDialog.SaveAddFilterItem(strFilter, "CSV (Comma delimited)", "*.csv")
    strPath = modSaveDialog.FileDialogSave1("", strPath, "Save CSV File", strFilter)
    
    ' Save template to CSV
    Call modDictionary.SaveAsCSV(strPath)
    Call modEvents.StartEvents
End Sub
Sub TemplateEditor_Click(control As IRibbonControl)
    Call modDictionary.OpenEditor
    Call modEvents.StartEvents
End Sub
Sub getInfo_Click(control As IRibbonControl)
    MsgBox "'SOAP Notes' was created by EszopiCoder, PharmD Student." & vbNewLine & _
"Open Source (https://github.com/EszopiCoder/word-soap-notes)" & vbNewLine & _
        "Please report bugs and send suggestions to pharm.coder@gmail.com", vbInformation
    Call modEvents.StartEvents
End Sub
