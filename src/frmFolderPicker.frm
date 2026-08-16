VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFolderPicker 
   Caption         =   "FolderPicker"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15015
   OleObjectBlob   =   "frmFolderPicker.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFolderPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' GLOBAL VARIABLES & CONSTANTS / GLOBALE VARIABLEN & KONSTANTEN
' =========================================================================
Private selectedFolder As Outlook.folder

Private Const APP_NAME As String = "OutlookFolderPicker"
Private Const SECTION_NAME As String = "History"

' =========================================================================
' FORM INITIALIZATION / FORMULAR-INITIALISIERUNG
' =========================================================================
Private Sub UserForm_Initialize()
    On Error GoTo EH
    
    ' 1. Apply global theme & title / Globales Theme & Titel setzen
    modUITheme.ApplyFormTheme Me
    Me.Caption = "FolderPicker"
    
    ' 2. Check if cache exists; build once if empty / Prüfen ob Cache existiert
    If gAllFolders Is Nothing Then
        RefreshFolderCache
    End If
    
    ' 3. Configure TreeView visual node lines (+/-) / TreeView-Linien konfigurieren
    With tvFolders
        .Nodes.Clear
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
    End With
    
    ' 4. Populate display from cache & history / Anzeige aus Cache & Historie aufbauen
    PopulateTreeViewFromCache ""
    RefreshHistoryListBox
    
    ' 5. Set search filter placeholder / Suchfilter-Platzhalter setzen
    txtFilter.Text = "Type to filter and press Enter"
    txtFilter.ForeColor = modUITheme.COLOR_FONT_MUTED
    
    ' 6. Initialize button styles / Button-Styles und Beschriftungen setzen
    SetupAllButtons
    
    ' 7. Check execution mode / Modus prüfen (E-Mails ausgewählt vs. Navigation)
    If IsGoToOnlyMode Then
        lblBtnMove.Enabled = False
        lblBtnCopy.Enabled = False
        lblBtnMove.BackColor = RGB(200, 200, 200)
        lblBtnCopy.BackColor = RGB(200, 200, 200)
    End If
    
    Exit Sub

EH:
    MsgBox "Error loading folder structure / Fehler beim Laden der Ordnerstruktur: " & Err.Description, vbCritical
End Sub

' =========================================================================
' BUTTON ACTIONS & EVENTS / BUTTON-AKTIONEN & EVENTS
' =========================================================================
Private Sub lblBtnMove_Click()
    ProcessEmails True
End Sub

Private Sub lblBtnCopy_Click()
    ProcessEmails False
End Sub

Private Sub lblBtnGoTo_Click()
    If selectedFolder Is Nothing Then
        MsgBox "Please select a folder first. / Bitte wählen Sie zuerst einen Ordner aus.", vbExclamation
        Exit Sub
    End If

    ' Check if target folder is already active / Prüfen, ob Ordner bereits aktiv ist
    On Error Resume Next
    Dim activeFolder As Outlook.folder
    Set activeFolder = Application.ActiveExplorer.currentFolder
    On Error GoTo 0

    If Not activeFolder Is Nothing Then
        If activeFolder.entryID = selectedFolder.entryID Then
            MsgBox "You are already in this folder. / Sie befinden sich bereits in diesem Ordner.", _
                   vbInformation, "FolderPicker"
            Exit Sub
        End If
    End If

    AddFolderToHistory selectedFolder
    Set gSelectedFolder = selectedFolder
    
    Unload Me
    Call AfterFolderPickerClosed
End Sub

Private Sub lblBtnNew_Click()
    On Error GoTo EH
    If selectedFolder Is Nothing Then
        MsgBox "Please select a parent folder first. / Bitte wählen Sie zuerst einen übergeordneten Ordner aus.", vbExclamation
        Exit Sub
    End If
    
    Dim newFolderName As String
    newFolderName = InputBox("Name of new subfolder: / Name des neuen Unterordners:", "New Folder / Neuer Ordner")
    If Trim(newFolderName) = "" Then Exit Sub
    
    Dim newFolder As Outlook.folder
    Set newFolder = selectedFolder.Folders.Add(newFolderName)
    
    RefreshFolderCache
    PopulateTreeViewFromCache txtFilter.Text
    
    MsgBox "Folder created successfully. / Ordner '" & newFolderName & "' erfolgreich erstellt.", vbInformation
    Exit Sub
EH:
    MsgBox "Error creating folder / Fehler beim Erstellen des Ordners: " & Err.Description, vbCritical
End Sub

Private Sub lblBtnRefresh_Click()
    On Error GoTo EH
    Me.MousePointer = fmMousePointerHourGlass
    
    RefreshFolderCache
    PopulateTreeViewFromCache txtFilter.Text
    
    Me.MousePointer = fmMousePointerDefault
    Exit Sub
EH:
    Me.MousePointer = fmMousePointerDefault
    MsgBox "Error during refresh / Fehler beim Aktualisieren: " & Err.Description, vbCritical
End Sub

Private Sub lblBtnSettings_Click()
    frmSettings.Show
    RefreshFolderCache
    PopulateTreeViewFromCache txtFilter.Text
End Sub

Private Sub lblBtnHelp_Click()
    frmAbout.Show vbModal
End Sub

Private Sub lblBtnCancel_Click()
    Unload Me
End Sub

' =========================================================================
' TREEVIEW & FILTER PROCESSING / TREEVIEW & FILTER-VERARBEITUNG
' =========================================================================
Public Sub PopulateTreeViewFromCache(ByVal filterText As String)
    On Error GoTo EH
    
    Dim Node As MSComctlLib.Node
    Dim parentNode As MSComctlLib.Node
    Dim fData As Collection
    Dim folder As Outlook.folder
    Dim level As Long
    
    Dim displayName As String, storeName As String, rootFolderName As String
    Dim isStoreActive As Boolean
    Dim startExpand As Long, endExpand As Long
    Dim allowedRootFolders As String
    Dim isVisible As Boolean
    Dim query As String
    
    query = LCase(Trim(filterText))
    If query = "type to filter and press enter" Then query = ""
    
    tvFolders.Nodes.Clear
    tvFolders.Sorted = True
    
    If gAllFolders Is Nothing Then Exit Sub
    
    For Each fData In gAllFolders
        Set folder = fData(1)
        level = fData(2)
        
        displayName = folder.Name
        storeName = folder.store.displayName
        
        isStoreActive = CBool(GetSetting(APP_NAME, storeName, "Active", "True"))
        
        If isStoreActive Then
            startExpand = Val(GetSetting(APP_NAME, storeName, "StartExpandLevel", "1"))
            endExpand = Val(GetSetting(APP_NAME, storeName, "EndExpandLevel", "3"))
            allowedRootFolders = GetSetting(APP_NAME, storeName, "AllowedRootFolders", "")
            
            isVisible = True
            
            ' Levels & Root check / Ebenen & Hauptordner prüfen
            If level = 0 Then
                If startExpand > 0 Then isVisible = False
            ElseIf level = 1 Then
                If allowedRootFolders <> "" Then
                    If InStr(1, ";" & allowedRootFolders & ";", ";" & displayName & ";", vbTextCompare) = 0 Then
                        isVisible = False
                    End If
                End If
            Else
                rootFolderName = GetRootFolderOfLevel1(folder)
                If allowedRootFolders <> "" And rootFolderName <> "" Then
                    If InStr(1, ";" & allowedRootFolders & ";", ";" & rootFolderName & ";", vbTextCompare) = 0 Then
                        isVisible = False
                    End If
                End If
            End If
            
            ' Search filter check / Suchfilter prüfen
            If isVisible And query <> "" Then
                If InStr(1, LCase(displayName), query, vbTextCompare) = 0 Then
                    isVisible = False
                End If
            End If
            
            ' Add node to TreeView / Knoten hinzufügen
            If isVisible Then
                Set Node = Nothing
                
                ' Flat display when searching / Flache Ansicht bei Suchfiltern
                If query <> "" Then
                    Dim fullPathDisplay As String
                    fullPathDisplay = displayName & "  (" & folder.FolderPath & ")"
                    
                    On Error Resume Next
                    Set Node = tvFolders.Nodes.Add(, , "ID_" & folder.entryID, fullPathDisplay)
                    On Error GoTo EH
                Else
                    ' Standard tree hierarchy / Standard Baum-Hierarchie
                    Set parentNode = Nothing
                    
                    If level > startExpand Then
                        On Error Resume Next
                        If Not folder.Parent Is Nothing Then
                            Set parentNode = tvFolders.Nodes("ID_" & folder.Parent.entryID)
                        End If
                        On Error GoTo EH
                    End If
                    
                    On Error Resume Next
                    If parentNode Is Nothing Then
                        Set Node = tvFolders.Nodes.Add(, , "ID_" & folder.entryID, displayName)
                    Else
                        Set Node = tvFolders.Nodes.Add(parentNode.Key, tvwChild, "ID_" & folder.entryID, displayName)
                    End If
                    On Error GoTo EH
                    
                    ' Expansion state / Aufklapp-Steuerung
                    If Not Node Is Nothing Then
                        If level >= startExpand And level < endExpand Then
                            Node.Expanded = True
                        Else
                            Node.Expanded = False
                        End If
                    End If
                End If
            End If
        End If
    Next fData
    
    ' Sort nodes / Knoten sortieren
    For Each Node In tvFolders.Nodes
        Node.Sorted = True
    Next Node
    
    Exit Sub

EH:
    Resume Next
End Sub

Private Function GetRootFolderOfLevel1(fld As Outlook.folder) As String
    On Error Resume Next
    GetRootFolderOfLevel1 = ""
    
    If fld Is Nothing Then Exit Function
    
    Dim pathParts() As String
    pathParts = Split(fld.FolderPath, "\")
    
    If UBound(pathParts) >= 3 Then
        GetRootFolderOfLevel1 = pathParts(3)
    End If
    On Error GoTo 0
End Function

Private Sub tvFolders_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim cleanEntryID As String
    
    If Left(Node.Key, 3) = "ID_" Then
        cleanEntryID = Mid(Node.Key, 4)
    Else
        cleanEntryID = Node.Key
    End If
    
    Set selectedFolder = Application.GetNamespace("MAPI").GetFolderFromID(cleanEntryID)
    On Error GoTo 0
End Sub

Private Sub tvFolders_DblClick()
    If selectedFolder Is Nothing Then Exit Sub
    If Not IsGoToOnlyMode Then ProcessEmails True
End Sub

Private Sub txtFilter_Enter()
    If txtFilter.ForeColor = modUITheme.COLOR_FONT_MUTED Then
        txtFilter.Text = ""
        txtFilter.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        PopulateTreeViewFromCache txtFilter.Text
        KeyCode = 0
    End If
End Sub

' =========================================================================
' MAIL PROCESSING ENGINE / E-MAIL VERARBEITUNG
' =========================================================================
Private Sub ProcessEmails(ByVal moveAction As Boolean)
    If IsGoToOnlyMode Then
        MsgBox "No emails selected. Use 'Go To' button. / Keine E-Mails ausgewählt. Verwenden Sie die 'Gehe zu'-Schaltfläche.", vbInformation
        Exit Sub
    End If

    If selectedFolder Is Nothing Then
        MsgBox "Please select a target folder. / Bitte wählen Sie einen Zielordner aus.", vbExclamation, "Sort / Sortieren"
        Exit Sub
    End If

    ' Check if source and target folders are identical / Quell- und Zielordner identisch?
    If SelectedItems.count > 0 Then
        Dim currentFolder As Outlook.folder
        On Error Resume Next
        Set currentFolder = SelectedItems.item(1).Parent
        On Error GoTo 0
        
        If Not currentFolder Is Nothing Then
            If currentFolder.entryID = selectedFolder.entryID Then
                MsgBox "Source and target folders are identical. Action cancelled. / " & vbCrLf & _
                       "Quell- und Zielordner sind identisch. Aktion abgebrochen.", _
                       vbExclamation, "FolderPicker"
                Exit Sub
            End If
        End If
    End If

    AddFolderToHistory selectedFolder
    
    On Error GoTo ErrHandler
    Dim itm As Object
    Dim i As Long
    
    For i = SelectedItems.count To 1 Step -1
        Set itm = SelectedItems.item(i)
        
        If TypeOf itm Is Outlook.MailItem Then
            If moveAction Then
                itm.Move selectedFolder
            Else
                Dim newMail As Outlook.MailItem
                Set newMail = itm.Copy
                newMail.Move selectedFolder
            End If
        End If
    Next i
    
    Unload Me
    Exit Sub
    
ErrHandler:
    MsgBox "Error processing emails / Fehler beim Verarbeiten der Mails: " & Err.Description, vbCritical
End Sub

' =========================================================================
' HISTORY MANAGEMENT (REGISTRY) / HISTORIE VERWALTUNG
' =========================================================================
Private Sub AddFolderToHistory(folder As Outlook.folder)
    If folder Is Nothing Then Exit Sub
    
    Dim history As Collection
    Set history = LoadHistoryCollection()
    
    Dim i As Long
    For i = history.count To 1 Step -1
        If Split(history(i), "|")(0) = folder.entryID Then history.Remove i
    Next i
    
    Dim newItem As String
    newItem = folder.entryID & "|" & folder.FolderPath
    
    If history.count = 0 Then
        history.Add newItem
    Else
        history.Add newItem, Before:=1
    End If
    
    While history.count > 20
        history.Remove history.count
    Wend
    
    SaveHistoryCollection history
    RefreshHistoryListBox
End Sub

Private Function LoadHistoryCollection() As Collection
    Dim coll As New Collection
    Dim rawData As String
    rawData = GetSetting(APP_NAME, SECTION_NAME, "RecentFolders", "")
    
    If rawData <> "" Then
        Dim items() As String, v As Variant
        items = Split(rawData, ";;")
        For Each v In items
            If Trim(v) <> "" Then coll.Add v
        Next v
    End If
    
    Set LoadHistoryCollection = coll
End Function

Private Sub SaveHistoryCollection(coll As Collection)
    Dim rawData As String, i As Long
    For i = 1 To coll.count
        rawData = rawData & coll(i) & ";;"
    Next i
    If Len(rawData) > 2 Then rawData = Left(rawData, Len(rawData) - 2)
    SaveSetting APP_NAME, SECTION_NAME, "RecentFolders", rawData
End Sub

Private Sub RefreshHistoryListBox()
    lstHistory.Clear
    Dim coll As Collection
    Set coll = LoadHistoryCollection()
    
    Dim item As Variant, parts() As String, fullPath As String, pathParts() As String
    Dim displayFolder As String
    
    For Each item In coll
        parts = Split(item, "|")
        If UBound(parts) >= 1 Then
            fullPath = Replace(parts(1), "%2F", "/")
            fullPath = Replace(fullPath, "%2f", "/")
            fullPath = Replace(fullPath, "%20", " ")
            
            If InStr(fullPath, "\") > 0 Then
                pathParts = Split(fullPath, "\")
                displayFolder = pathParts(UBound(pathParts))
            Else
                displayFolder = fullPath
            End If
            
            lstHistory.AddItem displayFolder
        End If
    Next item
End Sub

Private Sub lstHistory_Click()
    If lstHistory.ListIndex = -1 Then Exit Sub
    Dim coll As Collection
    Set coll = LoadHistoryCollection()
    
    Dim selectedData As String, entryID As String
    selectedData = coll(lstHistory.ListIndex + 1)
    entryID = Split(selectedData, "|")(0)
    
    On Error Resume Next
    Set selectedFolder = Application.GetNamespace("MAPI").GetFolderFromID(entryID)
    On Error GoTo 0
End Sub

Private Sub lstHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedFolder Is Nothing Then Exit Sub
    If IsGoToOnlyMode Then
        lblBtnGoTo_Click
    Else
        ProcessEmails True
    End If
End Sub

' =========================================================================
' STYLING & HOVER EFFECTS / DESIGNEINSTELLUNGEN & HOVER-EFFEKTE
' =========================================================================
Private Sub SetupAllButtons()
    lblBtnMove.Caption = "Move"
    modUITheme.StyleLabelButton lblBtnMove, True
    
    lblBtnGoTo.Caption = "GoTo"
    modUITheme.StyleLabelButton lblBtnGoTo, False
    
    lblBtnCopy.Caption = "Copy"
    modUITheme.StyleLabelButton lblBtnCopy, False
    
    lblBtnNew.Caption = "New"
    modUITheme.StyleLabelButton lblBtnNew, False
    
    lblBtnSettings.Caption = "Settings"
    modUITheme.StyleLabelButton lblBtnSettings, False
    
    lblBtnRefresh.Caption = "Refresh"
    modUITheme.StyleLabelButton lblBtnRefresh, False
    
    lblBtnCancel.Caption = "Cancel"
    modUITheme.StyleLabelButton lblBtnCancel, False
    
    lblBtnHelp.Caption = "About"
    modUITheme.StyleLabelButton lblBtnHelp, False
End Sub

Private Sub lblBtnMove_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not IsGoToOnlyMode Then
        SetupAllButtons
        modUITheme.ApplyButtonHover lblBtnMove, isPrimary:=True
    End If
End Sub

Private Sub lblBtnGoTo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnGoTo
End Sub

Private Sub lblBtnCopy_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not IsGoToOnlyMode Then
        SetupAllButtons
        modUITheme.ApplyButtonHover lblBtnCopy
    End If
End Sub

Private Sub lblBtnNew_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnNew
End Sub

Private Sub lblBtnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnCancel, isCancel:=True
End Sub

Private Sub lblBtnSettings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnSettings, isNeutral:=True
End Sub

Private Sub lblBtnRefresh_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnRefresh, isNeutral:=True
End Sub

Private Sub lblBtnHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnHelp, isNeutral:=True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
End Sub


