Attribute VB_Name = "modUITheme"
Option Explicit

' =========================================================================
' CENTRAL UI THEME & STYLING MODULE / ZENTRALES DESIGN- & STYLING-MODUL
' =========================================================================

' Central Color Constants / Zentrale Farbkonstanten
Public Const COLOR_FORM_BG As Long = 16119285       ' RGB(245, 245, 245) - Light Gray Form Background
Public Const COLOR_FONT_DARK As Long = 2631720      ' RGB(40, 40, 40)   - Primary Text Color
Public Const COLOR_FONT_MUTED As Long = 6579300     ' RGB(100, 100, 100) - Secondary/Version Text
Public Const COLOR_FONT_WHITE As Long = 16777215    ' RGB(255, 255, 255) - White Text

' Primary Actions (Move, Save) - Classic Subtle Blue / Primäre Aktionen - Dezent Blau
Public Const COLOR_PRIMARY As Long = 13924352       ' RGB(0, 120, 212)  - Primary Accent Blue
Public Const COLOR_PRIMARY_HOVER As Long = 11822080 ' RGB(0, 100, 180)  - Subtle Darker Blue Hover
Public Const COLOR_PRIMARY_BORDER As Long = 10510336 ' RGB(0, 80, 160)   - Primary Dark Border

' Standard Buttons / Standard Buttons
Public Const COLOR_BTN_BG As Long = 16777215        ' RGB(255, 255, 255) - Standard White Button
Public Const COLOR_BTN_HOVER As Long = 16443622     ' RGB(230, 240, 250) - Soft Blue Hover
Public Const COLOR_BTN_BORDER As Long = 13816530    ' RGB(210, 210, 210) - Light Border
Public Const COLOR_BTN_HOVER_BORDER As Long = 14464150 ' RGB(150, 180, 220) - Hover Border

' Cancel & Neutral Buttons / Abbrechen & Neutrale Buttons
Public Const COLOR_CANCEL_HOVER As Long = 15132394  ' RGB(250, 230, 230) - Soft Red Hover
Public Const COLOR_CANCEL_BORDER As Long = 10526940 ' RGB(220, 160, 160) - Cancel Border

Public Const COLOR_NEUTRAL_HOVER As Long = 15461355 ' RGB(235, 235, 235) - Soft Neutral Gray Hover
Public Const COLOR_NEUTRAL_BORDER As Long = 11842740 ' RGB(180, 180, 180) - Neutral Border

Public Const FONT_NAME_DEFAULT As String = "Segoe UI"

' Apply global theme to a UserForm and its controls / Globales Theme auf UserForm anwenden
Public Sub ApplyFormTheme(ByVal frm As MSForms.UserForm)
    frm.BackColor = COLOR_FORM_BG
    
    Dim ctrl As MSForms.Control
    For Each ctrl In frm.Controls
        On Error Resume Next
        ctrl.Font.Name = FONT_NAME_DEFAULT
        ctrl.Font.Size = 9.5
        
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ListBox" Or TypeName(ctrl) = "ComboBox" Then
            ctrl.SpecialEffect = fmSpecialEffectFlat
            ctrl.BorderStyle = fmBorderStyleSingle
            ctrl.BorderColor = RGB(200, 200, 200)
            ctrl.BackColor = RGB(255, 255, 255)
        End If
        On Error GoTo 0
    Next ctrl
End Sub

' Format a label to act as a styled button / Formatiert ein Label als Button
Public Sub StyleLabelButton(ByVal lbl As MSForms.Label, Optional ByVal isPrimary As Boolean = False)
    With lbl
        .TextAlign = fmTextAlignCenter
        .Font.Name = FONT_NAME_DEFAULT
        .Font.Size = 9.5
        .BorderStyle = fmBorderStyleSingle
        
        If isPrimary Then
            .BackColor = COLOR_PRIMARY
            .ForeColor = COLOR_FONT_WHITE
            .BorderColor = COLOR_PRIMARY_BORDER
            .Font.Bold = True
        Else
            .BackColor = COLOR_BTN_BG
            .ForeColor = COLOR_FONT_DARK
            .BorderColor = COLOR_BTN_BORDER
            .Font.Bold = False
        End If
    End With
End Sub

' Apply hover effects on label buttons / Hover-Effekt auf Label-Buttons anwenden
Public Sub ApplyButtonHover(ByVal lbl As MSForms.Label, Optional ByVal isPrimary As Boolean = False, _
                             Optional ByVal isCancel As Boolean = False, Optional ByVal isNeutral As Boolean = False)
    With lbl
        If isPrimary Then
            .BackColor = COLOR_PRIMARY_HOVER
            .BorderColor = COLOR_PRIMARY_BORDER
        ElseIf isCancel Then
            .BackColor = COLOR_CANCEL_HOVER
            .BorderColor = COLOR_CANCEL_BORDER
        ElseIf isNeutral Then
            .BackColor = COLOR_NEUTRAL_HOVER
            .BorderColor = COLOR_NEUTRAL_BORDER
        Else
            .BackColor = COLOR_BTN_HOVER
            .BorderColor = COLOR_BTN_HOVER_BORDER
        End If
    End With
End Sub

