VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl CodeBox 
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   3240
   ScaleWidth      =   3840
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   3525
      Top             =   3720
   End
   Begin VB.PictureBox lineNumPanel 
      AutoRedraw      =   -1  'True
      Height          =   3240
      Left            =   0
      ScaleHeight     =   3180
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   0
      Width           =   3810
      Begin RichTextLib.RichTextBox rtf1 
         Height          =   3195
         Left            =   735
         TabIndex        =   1
         Top             =   0
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   5636
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"UserControl1.ctx":0000
      End
   End
End
Attribute VB_Name = "CodeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' REQUIRED API CALL
    Private Declare Function SendMessage _
            Lib "user32" Alias "SendMessageA" ( _
            ByVal hwnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            lParam As Any) As Long

'Default Property Values:
    Const m_def_LineNumber_BackColor = &H8000000F
    Const m_def_LineNumber_ForeColor = &H80000012
    Const m_def_LineNumber_PanelWidth = 500

'Property Variables:
    Dim m_LineNumber_BackColor As OLE_COLOR
    Dim m_LineNumber_ForeColor As OLE_COLOR
    Dim m_LineNumber_PanelWidth As Integer

'Event Declarations:
    Event Click() 'MappingInfo=rtf1,rtf1,-1,Click
    Event DblClick() 'MappingInfo=rtf1,rtf1,-1,DblClick
    Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtf1,rtf1,-1,KeyDown
    Event KeyPress(KeyAscii As Integer) 'MappingInfo=rtf1,rtf1,-1,KeyPress
    Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtf1,rtf1,-1,KeyUp
    Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtf1,rtf1,-1,MouseDown
    Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtf1,rtf1,-1,MouseMove
    Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtf1,rtf1,-1,MouseUp
    Event Resize()
    
    

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtf1,rtf1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
    
    BackColor = rtf1.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    rtf1.BackColor() = New_BackColor
    PropertyChanged "BackColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtf1,rtf1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    
    Enabled = rtf1.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    rtf1.Enabled() = New_Enabled
    PropertyChanged "Enabled"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtf1,rtf1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    
    Set Font = rtf1.Font

End Property

Public Property Set Font(ByVal New_Font As Font)
    
    Set rtf1.Font = New_Font
    
    rtf1.SelStart = 0
    rtf1.SelLength = Len(rtf1.Text)
    
    rtf1.SelFontName = New_Font.Name
    rtf1.SelFontSize = New_Font.Size
    
    PropertyChanged "Font"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtf1,rtf1,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    
    BorderStyle = rtf1.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    
    rtf1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtf1,rtf1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
    
    rtf1.Refresh

End Sub







Private Sub lineNumPanel_Click()

End Sub

Private Sub rtf1_Change()
    
    ' REPRINT LINE NUMBERS
    Dim iCurrIndex As Long
    iCurrIndex = 1 + SendMessage(rtf1.hwnd, &HCE, 0, 0)
    PrintLineNumbers iCurrIndex
    
End Sub

Private Sub rtf1_Click()
    
    RaiseEvent Click

End Sub

Private Sub rtf1_DblClick()
    
    RaiseEvent DblClick

End Sub

Private Sub rtf1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub rtf1_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub rtf1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub rtf1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub rtf1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub rtf1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&h8000000f&
Public Property Get LineNumber_BackColor() As OLE_COLOR
Attribute LineNumber_BackColor.VB_Description = "Backcolor of the area where the line numbers are listed."
    
    LineNumber_BackColor = m_LineNumber_BackColor

End Property

Public Property Let LineNumber_BackColor(ByVal New_LineNumber_BackColor As OLE_COLOR)
    
    m_LineNumber_BackColor = New_LineNumber_BackColor
    lineNumPanel.BackColor = New_LineNumber_BackColor
    PropertyChanged "LineNumber_BackColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&h8000000f&
Public Property Get LineNumber_ForeColor() As OLE_COLOR
Attribute LineNumber_ForeColor.VB_Description = "Text Color of the line numbers."
    
    LineNumber_ForeColor = m_LineNumber_ForeColor

End Property

Public Property Let LineNumber_ForeColor(ByVal New_LineNumber_ForeColor As OLE_COLOR)
    
    m_LineNumber_ForeColor = New_LineNumber_ForeColor
    lineNumPanel.ForeColor = New_LineNumber_ForeColor
    lineNumPanel.Refresh
    
    PropertyChanged "LineNumber_ForeColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,750
Public Property Get LineNumber_PanelWidth() As Integer
    
    LineNumber_PanelWidth = m_LineNumber_PanelWidth
    
End Property

Public Property Let LineNumber_PanelWidth(ByVal New_LineNumber_PanelWidth As Integer)
    
    m_LineNumber_PanelWidth = New_LineNumber_PanelWidth
    PropertyChanged "LineNumber_PanelWidth"
    UserControl_Resize

End Property

Private Sub Timer1_Timer()

    Static iPrevIndex As Long
    Dim iCurrIndex As Long

    ' GET CURRENT "1st line in view"
    iCurrIndex = 1 + SendMessage(rtf1.hwnd, &HCE, 0, 0)
    
    ' IF IT IS DIFFERENT THAT LAST TIME, THEN RE-PRINT
    ' LINE NUMBERS STARTING WITH THIS NEW NUMBER
    If iCurrIndex <> iPrevIndex Then
        
        ' SET THE PREV AND CURRENT TO MATCH
        iPrevIndex = iCurrIndex
        
        ' REPRINT LINE NUMBERS
        PrintLineNumbers iCurrIndex
        
    End If
    
End Sub

Private Sub PrintLineNumbers(StartingLineNumber As Long)
            
        iCurrIndex = StartingLineNumber
        
        With lineNumPanel
        ' SET FONT TO SAME AS THAT OF RTF
        ' SO LINE NUMBERS LINE UP PROPERLY
            .Font.Size = rtf1.Font.Size
            .Font.Name = rtf1.Font.Name
            .Font.Bold = rtf1.Font.Bold
            .Font.Italic = rtf1.Font.Italic
        
        ' CLEAR LAST NUMBERS PRINTED
            .Cls
        
        ' SET Y-AXIS
            .CurrentY = 0
        
        End With
        
        ' PRINT LINE NUMBERS
        Do Until lineNumPanel.CurrentY > lineNumPanel.Height
        
            ' LEFT MARGIN
            lineNumPanel.CurrentX = 30
            lineNumPanel.Print Right$("0000" & Trim$(Str$(iCurrIndex)), 4)
            iCurrIndex = iCurrIndex + 1
            
            If iCurrIndex > (1 + rtf1.GetLineFromChar(Len(rtf1.Text))) Then Exit Do
            
        Loop

    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    m_LineNumber_BackColor = m_def_LineNumber_BackColor
    m_LineNumber_ForeColor = m_def_LineNumber_ForeColor
    m_LineNumber_PanelWidth = m_def_LineNumber_PanelWidth

End Sub



'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    rtf1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    rtf1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set rtf1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    rtf1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_LineNumber_BackColor = PropBag.ReadProperty("LineNumber_BackColor", m_def_LineNumber_BackColor)
    m_LineNumber_ForeColor = PropBag.ReadProperty("LineNumber_ForeColor", m_def_LineNumber_ForeColor)
    m_LineNumber_PanelWidth = PropBag.ReadProperty("LineNumber_PanelWidth", m_def_LineNumber_PanelWidth)
    lineNumPanel.BackColor = m_LineNumber_BackColor
    lineNumPanel.ForeColor = m_LineNumber_ForeColor
    
End Sub

Private Sub UserControl_Resize()

    ' SIZE PANEL
    lineNumPanel.Width = ScaleWidth
    lineNumPanel.Height = ScaleHeight
    
    ' SIZE TEXT BOX
    rtf1.Left = LineNumber_PanelWidth
    rtf1.Width = lineNumPanel.Width - rtf1.Left - 60
    rtf1.Height = lineNumPanel.Height - 60
    
    ' REPRINT LINE NUMBERS
    Dim iCurrIndex As Long
    iCurrIndex = 1 + SendMessage(rtf1.hwnd, &HCE, 0, 0)
    PrintLineNumbers iCurrIndex
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", rtf1.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", rtf1.Enabled, True)
    Call PropBag.WriteProperty("Font", rtf1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", rtf1.BorderStyle, 0)
    Call PropBag.WriteProperty("LineNumber_BackColor", m_LineNumber_BackColor, m_def_LineNumber_BackColor)
    Call PropBag.WriteProperty("LineNumber_ForeColor", m_LineNumber_ForeColor, m_def_LineNumber_ForeColor)
    Call PropBag.WriteProperty("LineNumber_PanelWidth", m_LineNumber_PanelWidth, m_def_LineNumber_PanelWidth)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function UpdateLineNumbers() As Long
    
    ' REPRINT LINE NUMBERS
    Dim iCurrIndex As Long
    iCurrIndex = 1 + SendMessage(rtf1.hwnd, &HCE, 0, 0)
    PrintLineNumbers iCurrIndex
    
End Function

