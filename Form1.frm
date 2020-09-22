VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Codebox Example"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin Project1.CodeBox CodeBox1 
      Height          =   3615
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineNumber_BackColor=   14737632
      LineNumber_ForeColor=   4194304
      LineNumber_PanelWidth=   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Resize()

    ' DEMONSTRATES HOW THE LINE NUMBERS
    ' ARE UPDATED WHEN CONTROL IS RESIZED
    On Error Resume Next
      CodeBox1.Width = ScaleWidth - 30
      CodeBox1.Height = ScaleHeight - 30 - CodeBox1.Top
      
        
End Sub
