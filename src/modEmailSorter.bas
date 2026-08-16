Attribute VB_Name = "modEmailSorter"
Option Explicit

' --- GLOBAL VARIABLES / GLOBALE VARIABLEN ---
Public SelectedItems As Selection         ' Holds currently selected emails / Hält die aktuell markierten E-Mails
Public gSelectedFolder As Outlook.folder    ' Stores the target folder / Nimmt den Zielordner auf
Public IsGoToOnlyMode As Boolean          ' Flag: True = No emails selected (folder navigation only) / Flag: True = Keine Mails gewählt (nur Ordner-Navigation)

' --- MAIN MACRO / HAUPT-MAKRO ---
Public Sub RunFolderPicker()
    Dim sel As Selection
    
    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    On Error GoTo 0

    ' Check if emails are selected / Prüfen, ob E-Mails markiert sind
    If sel Is Nothing Then
        IsGoToOnlyMode = True
        Set SelectedItems = Nothing
    ElseIf sel.count = 0 Then
        IsGoToOnlyMode = True
        Set SelectedItems = Nothing
    Else
        IsGoToOnlyMode = False
        Set SelectedItems = sel
    End If

    ' Open userform / Formular öffnen
    frmFolderPicker.Show
End Sub

' --- CALLED AFTER CLOSING FOR "GO TO FOLDER" / AUFRUF NACH DEM SCHLIESSEN FÜR "GEHE ZU ORDNER" ---
Public Sub AfterFolderPickerClosed()
    On Error GoTo EH

    If gSelectedFolder Is Nothing Then Exit Sub

    ' Display selected folder in main explorer window / Ausgewählten Ordner im Hauptfenster anzeigen
    Set Application.ActiveExplorer.currentFolder = gSelectedFolder

    Exit Sub

EH:
    MsgBox "Error switching folder / Fehler beim Wechseln des Ordners: " & Err.Description, vbCritical, "Navigate Folder / Ordner navigieren"
End Sub
