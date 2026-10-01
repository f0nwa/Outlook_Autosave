VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderSelectionForm 
   Caption         =   "Select Outlook folders"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "FolderSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FolderSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents btnSave As MSForms.CommandButton
Attribute btnSave.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1
Private WithEvents btnClear As MSForms.CommandButton
Attribute btnClear.VB_VarHelpID = -1
Private WithEvents btnAddCurrent As MSForms.CommandButton
Attribute btnAddCurrent.VB_VarHelpID = -1

Private lstFolders As MSForms.ListBox
Private loadedFolderKeys As Object
Private isSingleFolderPicker As Boolean
Private pickedFolderPath As String

Private Sub UserForm_Initialize()
    BuildUi
    LoadFolders
End Sub

Public Function PickSingleFolderPath() As String
    isSingleFolderPicker = True
    pickedFolderPath = ""
    ConfigureSingleFolderPicker
    Me.Show
    PickSingleFolderPath = pickedFolderPath
    Unload Me
End Function

Private Sub BuildUi()
    Dim lblTitle As MSForms.Label
    Dim lblHint As MSForms.Label
    
    Me.Caption = RuText(1042, 1099, 1073, 1086, 1088, 32, 1087, 1072, 1087, 1086, 1082, 32, 79, 117, 116, 108, 111, 111, 107)
    Me.Width = 610
    Me.Height = 455
    
    Set lblTitle = AddLabel("lblTitle", RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1080, 32, 1076, 1083, 1103, 32, 1072, 1074, 1090, 1086, 1084, 1072, 1090, 1080, 1095, 1077, 1089, 1082, 1086, 1075, 1086, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081), 18, 12, 520, 18)
    lblTitle.Font.Bold = True
    
    Set lblHint = AddLabel("lblHint", RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1080, 32, 1079, 1076, 1077, 1089, 1100, 32, 1080, 1083, 1080, 32, 1074, 1099, 1076, 1077, 1083, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1091, 32, 1074, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1080, 32, 1085, 1072, 1078, 1084, 1080, 1090, 1077, 32, 34, 1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1091, 1102, 34, 46), 18, 36, 560, 18)
    
    Set lstFolders = Me.Controls.Add("Forms.ListBox.1", "lstFolders", True)
    With lstFolders
        .Left = 18
        .Top = 62
        .Width = 548
        .Height = 310
        .ColumnCount = 2
        .ColumnWidths = "520 pt;0 pt"
        .ListStyle = fmListStyleOption
        .MultiSelect = fmMultiSelectMulti
        .IntegralHeight = False
    End With
    
    Set btnSave = AddButton("btnSave", RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100), 18, 388, 100, 26, RuText(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1099, 1077, 32, 1087, 1072, 1087, 1082, 1080, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1076, 1083, 1103, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46))
    Set btnClear = AddButton("btnClear", RuText(1054, 1095, 1080, 1089, 1090, 1080, 1090, 1100), 128, 388, 98, 26, RuText(1057, 1085, 1103, 1090, 1100, 32, 1074, 1099, 1073, 1086, 1088, 32, 1089, 1086, 32, 1074, 1089, 1077, 1093, 32, 1087, 1072, 1087, 1086, 1082, 32, 1072, 1074, 1090, 1086, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 46))
    Set btnAddCurrent = AddButton("btnAddCurrent", RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1091, 1102), 238, 388, 150, 26, RuText(1044, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1091, 1102, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1085, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107, 46))
    Set btnCancel = AddButton("btnCancel", RuText(1054, 1090, 1084, 1077, 1085, 1072), 468, 388, 98, 26, RuText(1047, 1072, 1082, 1088, 1099, 1090, 1100, 32, 1073, 1077, 1079, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1080, 1079, 1084, 1077, 1085, 1077, 1085, 1080, 1081, 46))
End Sub

Private Sub LoadFolders()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim inboxFolder As Outlook.MAPIFolder
    Dim selectedKeys As Object
    
    Set selectedKeys = LoadSelectedKeys()
    Set loadedFolderKeys = CreateObject("Scripting.Dictionary")
    Set ns = Outlook.Application.GetNamespace("MAPI")
    
    lstFolders.Clear
    
    For Each store In ns.Stores
        Set inboxFolder = Nothing
        
        On Error Resume Next
        Set inboxFolder = store.GetDefaultFolder(olFolderInbox)
        On Error GoTo 0

        If Not inboxFolder Is Nothing Then
            AddFolderToList inboxFolder, 0, selectedKeys
        End If
    Next store
    
    AddSavedFoldersToList
End Sub

Private Sub AddSavedFoldersToList()
    Dim savedFolders As String
    Dim records() As String
    Dim record As String
    Dim i As Long
    
    savedFolders = GetSavedAutoSaveFolderRecords()
    
    If Len(Trim(savedFolders)) = 0 Then
        Exit Sub
    End If
    
    records = Split(savedFolders, vbLf)
    
    For i = LBound(records) To UBound(records)
        record = Trim(records(i))
        
        If Len(record) > 0 Then
            AddSavedRecordToList record
        End If
    Next i
End Sub

Private Sub AddSavedRecordToList(ByVal record As String)
    Dim key As String
    Dim rowIndex As Long
    
    key = GetFolderRecordKey(record)
    
    If Len(key) = 0 Then
        Exit Sub
    End If
    
    For rowIndex = 0 To lstFolders.ListCount - 1
        If CStr(lstFolders.List(rowIndex, 1)) = record Then
            lstFolders.Selected(rowIndex) = True
            Exit Sub
        End If
    Next rowIndex
    
    If loadedFolderKeys.Exists(key) Then
        Exit Sub
    End If
    
    loadedFolderKeys(key) = True
    lstFolders.AddItem GetFolderPathFromRecord(record)
    rowIndex = lstFolders.ListCount - 1
    lstFolders.List(rowIndex, 1) = record
    lstFolders.Selected(rowIndex) = True
End Sub

Private Function AddSingleFolderToList(ByVal folder As Outlook.MAPIFolder, ByVal displayName As String, ByVal selectFolder As Boolean) As Boolean
    Dim record As String
    Dim key As String
    Dim rowIndex As Long
    
    If Not ShouldShowMailFolder(folder) Then
        AddSingleFolderToList = False
        Exit Function
    End If
    
    record = folder.StoreID & "|" & folder.EntryID & "|" & folder.FolderPath
    key = GetFolderRecordKey(record)
    
    For rowIndex = 0 To lstFolders.ListCount - 1
        If CStr(lstFolders.List(rowIndex, 1)) = record Then
            If selectFolder Then
                lstFolders.Selected(rowIndex) = True
            End If
            
            AddSingleFolderToList = True
            Exit Function
        End If
    Next rowIndex
    
    loadedFolderKeys(key) = True
    lstFolders.AddItem displayName
    rowIndex = lstFolders.ListCount - 1
    lstFolders.List(rowIndex, 1) = record
    
    If selectFolder Then
        lstFolders.Selected(rowIndex) = True
    End If
    
    AddSingleFolderToList = True
End Function

Private Sub AddFolderToList(ByVal folder As Outlook.MAPIFolder, ByVal level As Long, ByVal selectedKeys As Object)
    Dim childFolder As Outlook.MAPIFolder
    Dim record As String
    Dim key As String
    Dim rowIndex As Long
    
    If Not ShouldShowMailFolder(folder) Then
        Exit Sub
    End If
    
    record = folder.StoreID & "|" & folder.EntryID & "|" & folder.FolderPath
    key = GetFolderRecordKey(record)
    
    If Not loadedFolderKeys.Exists(key) Then
        loadedFolderKeys(key) = True
        
        lstFolders.AddItem GetFolderListDisplayName(folder, level)
        rowIndex = lstFolders.ListCount - 1
        lstFolders.List(rowIndex, 1) = record
        
        If selectedKeys.Exists(key) Then
            lstFolders.Selected(rowIndex) = True
        End If
    End If
    
    For Each childFolder In folder.Folders
        AddFolderToList childFolder, level + 1, selectedKeys
    Next childFolder
End Sub

Private Function GetFolderListDisplayName(ByVal folder As Outlook.MAPIFolder, ByVal level As Long) As String
    Dim displayName As String
    
    displayName = String(level * 2, " ") & folder.Name
    
    If level = 0 Then
        displayName = folder.Store.DisplayName & " \ " & displayName
    End If
    
    GetFolderListDisplayName = displayName
End Function

Private Function ShouldShowMailFolder(ByVal folder As Outlook.MAPIFolder) As Boolean
    Dim folderName As String
    
    On Error GoTo HideFolder
    
    If folder.DefaultItemType <> olMailItem Then
        ShouldShowMailFolder = False
        Exit Function
    End If
    
    folderName = LCase(folder.Name)
    
    If Left(folderName, 1) = "{" And Right(folderName, 1) = "}" Then
        ShouldShowMailFolder = False
        Exit Function
    End If
    
    Select Case folderName
        Case "recipient cache", "peoplecentricconversation buddies", "organizational contacts", "gal contacts", _
             "contacts", "contactos", "calendar", "notes", "tasks", "journal", _
             "outbound", "inbound", "feeds"
            ShouldShowMailFolder = False
        Case Else
            ShouldShowMailFolder = True
    End Select
    
    Exit Function
    
HideFolder:
    ShouldShowMailFolder = False
End Function

Private Function LoadSelectedKeys() As Object
    Dim selectedKeys As Object
    Dim records() As String
    Dim savedFolders As String
    Dim i As Long
    Dim key As String
    
    Set selectedKeys = CreateObject("Scripting.Dictionary")
    savedFolders = GetSavedAutoSaveFolderRecords()
    
    If Len(Trim(savedFolders)) > 0 Then
        records = Split(savedFolders, vbLf)
        
        For i = LBound(records) To UBound(records)
            key = GetFolderRecordKey(records(i))
            
            If Len(key) > 0 Then
                selectedKeys(key) = True
            End If
        Next i
    End If
    
    Set LoadSelectedKeys = selectedKeys
End Function

Private Function GetFolderRecordKey(ByVal record As String) As String
    Dim parts() As String
    
    If Len(Trim(record)) = 0 Then
        GetFolderRecordKey = ""
        Exit Function
    End If
    
    parts = Split(record, "|")
    
    If UBound(parts) >= 1 Then
        GetFolderRecordKey = parts(0) & "|" & parts(1)
    Else
        GetFolderRecordKey = ""
    End If
End Function

Private Function GetFolderPathFromRecord(ByVal record As String) As String
    Dim parts() As String
    
    parts = Split(record, "|")
    
    If UBound(parts) >= 2 And Len(Trim(parts(2))) > 0 Then
        GetFolderPathFromRecord = parts(2)
    Else
        GetFolderPathFromRecord = record
    End If
End Function

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    
    With AddLabel
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
    End With
End Function

Private Function AddButton(ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single, ByVal heightValue As Single, Optional ByVal tipText As String = "") As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    
    With AddButton
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .ControlTipText = tipText
    End With
End Function

Private Sub ConfigureSingleFolderPicker()
    Me.Caption = RuText(1042, 1099, 1073, 1086, 1088, 32, 1087, 1072, 1087, 1082, 1080, 32, 79, 117, 116, 108, 111, 111, 107)
    Me.Controls("lblTitle").Caption = RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1080, 1089, 1093, 1086, 1076, 1085, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107)
    Me.Controls("lblHint").Caption = RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1086, 1076, 1085, 1091, 32, 1087, 1072, 1087, 1082, 1091, 32, 1080, 32, 1085, 1072, 1078, 1084, 1080, 1090, 1077, 32, 34, 1042, 1099, 1073, 1088, 1072, 1090, 1100, 34, 46)
    btnSave.Caption = RuText(1042, 1099, 1073, 1088, 1072, 1090, 1100)
    btnSave.ControlTipText = RuText(1042, 1099, 1073, 1088, 1072, 1090, 1100, 32, 1101, 1090, 1091, 32, 1087, 1072, 1087, 1082, 1091, 32, 1076, 1083, 1103, 1087, 1088, 1072, 1074, 1080, 1083, 1072, 46)
    btnClear.Visible = False
    btnAddCurrent.Visible = False
    ClearFolderSelections
    lstFolders.MultiSelect = fmMultiSelectSingle
    lstFolders.ListIndex = -1
End Sub

Private Sub ClearFolderSelections()
    Dim i As Long
    
    For i = 0 To lstFolders.ListCount - 1
        lstFolders.Selected(i) = False
    Next i
End Sub

Private Sub btnSave_Click()
    Dim i As Long
    Dim result As String
    
    If isSingleFolderPicker Then
        SavePickedSingleFolder
        Exit Sub
    End If
    
    For i = 0 To lstFolders.ListCount - 1
        If lstFolders.Selected(i) Then
            result = result & CStr(lstFolders.List(i, 1)) & vbLf
        End If
    Next i
    
    SaveAutoSaveFolderRecords result
    MsgBox RuText(1053, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1086, 32, 1087, 1072, 1087, 1086, 1082, 32, 79, 117, 116, 108, 111, 111, 107, 58, 32) & CountSelectedFolderRecords(result), vbInformation
    Unload Me
End Sub

Private Sub SavePickedSingleFolder()
    If lstFolders.ListIndex < 0 Then
        MsgBox RuText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1091, 46), vbInformation
        Exit Sub
    End If
    
    pickedFolderPath = GetFolderPathFromRecord(CStr(lstFolders.List(lstFolders.ListIndex, 1)))
    Me.Hide
End Sub

Private Sub btnClear_Click()
    Dim i As Long
    
    For i = 0 To lstFolders.ListCount - 1
        lstFolders.Selected(i) = False
    Next i
End Sub

Private Sub btnAddCurrent_Click()
    On Error GoTo ErrorHandler
    
    Dim currentFolder As Outlook.MAPIFolder
    
    Set currentFolder = Outlook.Application.ActiveExplorer.CurrentFolder
    
    If currentFolder Is Nothing Then
        MsgBox RuText(1058, 1077, 1082, 1091, 1097, 1072, 1103, 32, 1087, 1072, 1087, 1082, 1072, 32, 79, 117, 116, 108, 111, 111, 107, 32, 1085, 1077, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1072, 46), vbInformation
        Exit Sub
    End If
    
    If AddSingleFolderToList(currentFolder, currentFolder.FolderPath, True) Then
        MsgBox RuText(1055, 1072, 1087, 1082, 1072, 32, 1076, 1086, 1073, 1072, 1074, 1083, 1077, 1085, 1072, 32, 1080, 32, 1074, 1099, 1073, 1088, 1072, 1085, 1072, 58) & vbCrLf & currentFolder.FolderPath, vbInformation
    Else
        MsgBox RuText(1058, 1077, 1082, 1091, 1097, 1072, 1103, 32, 1087, 1072, 1087, 1082, 1072, 32, 1085, 1077, 32, 1103, 1074, 1083, 1103, 1077, 1090, 1089, 1103, 32, 1087, 1086, 1095, 1090, 1086, 1074, 1086, 1081, 32, 1080, 32, 1085, 1077, 32, 1084, 1086, 1078, 1077, 1090, 32, 1080, 1089, 1087, 1086, 1083, 1100, 1079, 1086, 1074, 1072, 1090, 1100, 1089, 1103, 32, 1076, 1083, 1103, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1074, 1083, 1086, 1078, 1077, 1085, 1080, 1081, 46), vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox RuText(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1076, 1086, 1073, 1072, 1074, 1080, 1090, 1100, 32, 1090, 1077, 1082, 1091, 1097, 1091, 1102, 32, 1087, 1072, 1087, 1082, 1091, 32, 79, 117, 116, 108, 111, 111, 107, 58, 32) & Err.Description, vbCritical
End Sub

Private Sub btnCancel_Click()
    If isSingleFolderPicker Then
        pickedFolderPath = ""
        Me.Hide
    Else
        Unload Me
    End If
End Sub
