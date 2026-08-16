VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Mailsorter"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' FORM INITIALIZATION / FORMULAR-INITIALISIERUNG
' =========================================================================
Private Sub UserForm_Initialize()
    ' 1. Base theme & title / Basis-Theme & Titel
    modUITheme.ApplyFormTheme Me
    Me.Caption = "About MailSorter"
    
    ' 2. Captions & descriptions / Texte & Beschreibungen
    lblVersion.Caption = "Version 1.0.0 (Release: August 2026)"
    lblDescription.Caption = "A lightweight, efficient utility designed to streamline and organize Outlook mailboxes and folder structures." & vbCrLf & vbCrLf & _
                             "Thank you for using this tool! If you find it helpful, feel free to support further development or share your feedback."
    
    ' 3. Typography formatting / Typografie-Formatierung
    lblVersion.Font.Name = modUITheme.FONT_NAME_DEFAULT
    lblVersion.Font.Size = 9
    lblVersion.ForeColor = modUITheme.COLOR_FONT_MUTED
    
    lblDescription.Font.Name = modUITheme.FONT_NAME_DEFAULT
    lblDescription.Font.Size = 10
    lblDescription.ForeColor = modUITheme.COLOR_FONT_DARK
    
    ' 4. Initialize button styles / Button-Formatierung initialisieren
    SetupAllButtons
End Sub

' =========================================================================
' BUTTON ACTIONS / BUTTON-AKTIONEN
' =========================================================================
Private Sub lblBtnRepo_Click()
    On Error Resume Next
    CreateObject("WScript.Shell").Run "https://github.com/chivas03/Mailsorter/", 0, False
    On Error GoTo 0
End Sub

Private Sub lblBtnCoffee_Click()
    On Error Resume Next
    CreateObject("WScript.Shell").Run "https://paypal.me/chivas03", 0, False
    On Error GoTo 0
End Sub

Private Sub lblBtnFeedback_Click()
    Dim formUrl As String
    formUrl = "https://docs.google.com/forms/d/e/1FAIpQLSccCMIOdzib0A-t94U5vFhSl7rKdssKgOj-FARel-UteiEM-g/viewform"
    On Error Resume Next
    CreateObject("WScript.Shell").Run formUrl, 0, False
    On Error GoTo 0
End Sub

Private Sub lblBtnClose_Click()
    Unload Me
End Sub

' =========================================================================
' STYLING & HOVER EFFECTS / DESIGNEINSTELLUNGEN & HOVER-EFFEKTE
' =========================================================================
Private Sub SetupAllButtons()
    lblBtnRepo.Caption = "GitHub Repository"
    modUITheme.StyleLabelButton lblBtnRepo, False
    
    lblBtnCoffee.Caption = "Treat a friend to a coffee!"
    modUITheme.StyleLabelButton lblBtnCoffee, False
    
    lblBtnFeedback.Caption = "Share Feedback"
    modUITheme.StyleLabelButton lblBtnFeedback, False
    
    lblBtnClose.Caption = "Close"
    modUITheme.StyleLabelButton lblBtnClose, False
End Sub

Private Sub lblBtnRepo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnRepo
End Sub

Private Sub lblBtnCoffee_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnCoffee
End Sub

Private Sub lblBtnFeedback_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnFeedback
End Sub

Private Sub lblBtnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
    modUITheme.ApplyButtonHover lblBtnClose, isNeutral:=True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetupAllButtons
End Sub

