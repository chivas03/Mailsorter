VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' CONSTANTS & VARIABLES / KONSTANTEN & VARIABLEN
' =========================================================================
Private Const APP_NAME As String = "OutlookFolderPicker"
Private isLoading As Boolean
Private currentStoreName As String

' =========================================================================
' FORM INITIALIZATION & LISTBOX ALIGNMENT (frmSettings)
' =========================================================================
Private Sub UserForm_Initialize()
    On Error GoTo EH
    isLoading = True
    
    ' 1. Apply global theme & title / Globales Theme & Titel
    modUITheme.ApplyFormTheme Me
    Me.Caption = "Settings"
    
    ' 2. Correct ListBox sizing & layout / ListBox-Anzeige korrigieren
    With lstAccounts
        .IntegralHeight = True  ' Erlaubt sauberes Einrasten ganzer Zeilen
        .Height = 125
        .Top = 30
    End With
    
    With lstStores
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .IntegralHeight = True  ' Verhindert abgeschnittene Zeilen am Ende
        .Height = 125
        .Top = 30
    End With
    
    ' 3. Populate levels and initialize button styles / Ebene-ComboBoxen füllen & Buttons stylen
    InitLevelComboBoxes
    SetupAllButtons
    
    ' 4. Load mailboxes / Postfaecher laden
    Dim ns As Outlook.NameSpace
    Dim st As Outlook.store
    Set ns = Application.GetNamespace("MAPI")
    
    lstAccounts.Clear
    For Each st In ns.Stores
        lstAccounts.AddItem st.displayName
    Next st
    
    ' 5. Select first account by default / Erstes Konto standardmaessig laden
    If lstAccounts.ListCount > 0 Then
        lstAccounts.ListIndex = 0
        currentStoreName = lstAccounts.List(0)
        LoadStoreSettings currentStoreName
    End If
    
    isLoading = False
    Exit Sub

EH:
    isLoading = False
    MsgBox "Error loading settings / Fehler beim Laden der Einstellungen: " & Err.Description, vbCritical
End Sub

Private Sub InitLevelComboBoxes()
    Dim i As Long
    On Error Resume Next
    cmbStartLevel.Clear
    cmbEndLevel.Clear
    
    cmbStartLevel.Style = fmStyleDropDownList
    cmbEndLevel.Style = fmStyleDropDownList
    
    For i = 0 To 10
        cmbStartLevel.AddItem CStr(i)
        cmbEndLevel.AddItem CStr(i)
    Next i
    On Error GoTo 0
End Sub

' =========================================================================
' UI EVENTS & CONTROLS / BENUTZEROBERFLÄCHEN-EVENTS
' =========================================================================
Private Sub lstAccounts_Click()
    If isLoading Then Exit Sub
    If lstAccounts.ListIndex = -1 Then Exit Sub
    
    If currentStoreName <> "" Then
        SaveStoreSettings currentStoreName
    End If
    
    isLoading = True
    currentStoreName = lstAccounts.List(lstAccounts.ListIndex)
    LoadStoreSettings currentStoreName
    isLoading = False
End Sub

Private Sub chkStoreActive_Click()
    If isLoading Then Exit Sub
    
    Dim isActive As Boolean
    isActive = chkStoreActive.Value
    
    cmbStartLevel.Enabled = isActive
    cmbEndLevel.Enabled = isActive
    lstStores.Enabled = isActive
End Sub

' =========================================================================
' LOAD & SAVE LOGIC / SPEICHER- UND LADE-LOGIK
' =========================================================================
Private Sub LoadStoreSettings(ByVal storeName As String)
    On Error GoTo EH
    
    Dim isActive As Boolean
    isActive = CBool(GetSetting(APP_NAME, storeName, "Active", "True"))
    chkStoreActive.Value = isActive
    
    cmbStartLevel.Value = GetSetting(APP_NAME, storeName, "StartExpandLevel", "1")
    cmbEndLevel.Value = GetSetting(APP_NAME, storeName, "EndExpandLevel", "3")
    
    PopulateRootFolderList storeName
    
    cmbStartLevel.Enabled = isActive
    cmbEndLevel.Enabled = isActive
    lstStores.Enabled = isActive
    Exit Sub

EH:
    MsgBox "Error loading store settings / Fehler beim Laden der Postfach-Einstellungen: " & Err.Description, vbCritical
End Sub

Private Sub SaveStoreSettings(ByVal storeName As String)
    If storeName = "" Then Exit Sub
    
    SaveSetting APP_NAME, storeName, "Active", CStr(chkStoreActive.Value)
    SaveSetting APP_NAME, storeName, "StartExpandLevel", cmbStartLevel.Value
    SaveSetting APP_NAME, storeName, "EndExpandLevel", cmbEndLevel.Value
    
    Dim allowedFolders As String
    Dim i As Long
    
    allowedFolders = ""
    For i = 0 To lstStores.ListCount - 1
        If lstStores.Selected(i) Then
            allowedFolders = allowedFolders & lstStores.List(i) & ";"
        End If
    Next i
    
    If Len(allowedFolders) > 0 Then
        allowedFolders = Left(allowedFolders, Len(allowedFolders) - 1)
    End If
    
    SaveSetting APP_NAME, storeName, "AllowedRootFolders", allowedFolders
End Sub

Private Sub PopulateRootFolderList(ByVal storeName As String)
    On Error GoTo EH
    
    lstStores.Clear
    
    Dim ns As Outlook.NameSpace
    Dim st As Outlook.store
    Dim rootFld As Outlook.folder
    Dim subFld As Outlook.folder
    Dim allowedStr As String
    Dim allowedArray() As String
    Dim j As Long
    Dim isSelected As Boolean
    Dim hasSavedSettings As Boolean
    
    Set ns = Application.GetNamespace("MAPI")
    allowedStr = GetSetting(APP_NAME, storeName, "AllowedRootFolders", "___DEFAULT___")
    
    If allowedStr <> "___DEFAULT___" Then
        hasSavedSettings = True
        allowedArray = Split(allowedStr, ";")
    Else
        hasSavedSettings = False
    End If
    
    For Each st In ns.Stores
        If st.displayName = storeName Then
            Set rootFld = st.GetRootFolder
            Exit For
        End If
    Next st
    
    If rootFld Is Nothing Then Exit Sub
    
    For Each subFld In rootFld.Folders
        lstStores.AddItem subFld.Name
        
        isSelected = False
        If hasSavedSettings Then
            For j = LBound(allowedArray) To UBound(allowedArray)
                If LCase(Trim(allowedArray(j))) = LCase(Trim(subFld.Name)) Then
                    isSelected = True
                    Exit For
                End If
            Next j
        Else
            isSelected = True
        End If
        
        lstStores.Selected(lstStores.ListCount - 1) = isSelected
    Next subFld
    
    Exit Sub

EH:
    MsgBox "Error loading folder list / Fehler beim Laden der Ordnerliste: " & Err.Description, vbCritical
End Sub

' =========================================================================
' BUTTON CLICK EVENTS / ACTION-EVENTS DER BUTTONS
' =========================================================================
Private Sub lblBtnSave_Click()
    If currentStoreName <> "" Then
        SaveStoreSettings currentStoreName
    End If
    Unload Me
End Sub

Private Sub lblBtnCancel_Click()
    Unload Me
End Sub

Private Sub lblBtnHelp_Click()
    frmAbout.Show vbModal
End Sub

' =========================================================================
' STYLING & HOVER EFFECTS / DESIGNEINSTELLUNGEN & HOVER-EFFEKTE
' =========================================================================
Private Sub SetupAllButtons()
    lblBtnSave.Caption = "Save"
    modUITheme.StyleLabelButton lblBtnSave, True
    
    lblBtnCancel.Caption = "Cancel"
    modUITheme.StyleLabelButton lblBtnCancel, False
    
    lblBtnHelp.Caption = "About"
    modUITheme.StyleLabelButton lblBtnHelp, False
End Sub

Private Sub lblBtnSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnSave, isPrimary:=True
End Sub

Private Sub lblBtnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnCancel, isCancel:=True
End Sub

Private Sub lblBtnHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnHelp, isNeutral:=True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
End Sub
