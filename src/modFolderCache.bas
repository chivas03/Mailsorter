Attribute VB_Name = "modFolderCache"
Option Explicit

Public gAllFolders As Collection
Private Const APP_NAME As String = "OutlookFolderPicker"

Public Sub RefreshFolderCache()
    Set gAllFolders = New Collection
    
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    Dim store As Outlook.store
    Dim rootFolder As Outlook.folder
    
    ' ... (Dein bisheriger Code zum Laden der disabled stores bleibt gleich)

    For Each store In ns.Stores
        ' Prüfung auf "IsConnected" (Nur für Exchange Stores sinnvoll)
        ' Wenn der Store offline ist, überspringen wir ihn, um Outlook nicht einzufrieren
        If store.ExchangeStoreType = olExchangeMailbox Or store.ExchangeStoreType = olExchangePublicFolder Then
            If Not store.IsOpen Then
                ' Store ist nicht verbunden, wir überspringen ihn komplett
                GoTo NextStore
            End If
        End If

        ' Zugriff absichern
        On Error Resume Next
        Set rootFolder = store.GetRootFolder
        
        ' Wenn Err.Number <> 0, ist der Store vermutlich nicht erreichbar/gesperrt
        If Err.Number <> 0 Or rootFolder Is Nothing Then
            Err.Clear
            GoTo NextStore ' Direkt zum nächsten Store springen
        End If
        On Error GoTo 0
        
        ' Erst jetzt cachen, wenn der Zugriff erfolgreich war
        CacheSubFolders rootFolder, 0

NextStore:
        ' Fehler zurücksetzen und weiter
        On Error GoTo 0
    Next store
End Sub

Private Sub CacheSubFolders(ByVal parentFolder As Outlook.folder, ByVal level As Long)
    ' Read maximum folder depth limit / Maximale Verzeichnistiefe auslesen
    Dim maxDepth As Long
    maxDepth = Val(GetSetting(APP_NAME, "Settings", "MaxTreeDepth", "10"))
    If maxDepth < 10 Then maxDepth = 10
    
    ' Exit if current level exceeds limit / Abbrechen wenn Tiefe überschritten ist
    If level > maxDepth Then Exit Sub

    Dim folderData As Collection
    Set folderData = New Collection
    folderData.Add parentFolder
    folderData.Add level
    
    gAllFolders.Add folderData

    On Error Resume Next
    Dim subFoldersCount As Long
    subFoldersCount = parentFolder.Folders.count
    On Error GoTo 0

    ' Recursively process subfolders / Unterordner rekursiv verarbeiten
    If subFoldersCount > 0 Then
        Dim subFolder As Outlook.folder
        For Each subFolder In parentFolder.Folders
            CacheSubFolders subFolder, level + 1
        Next subFolder
    End If
End Sub
