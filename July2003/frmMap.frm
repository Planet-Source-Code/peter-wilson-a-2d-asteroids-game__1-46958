VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Map - Press TAB to toggle."
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleMode       =   0  'User
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictMap 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H80000008&
      Height          =   1815
      Left            =   60
      ScaleHeight     =   1815
      ScaleWidth      =   2805
      TabIndex        =   0
      Top             =   60
      Width           =   2805
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   705
         Left            =   930
         Top             =   390
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_blnMouseDown  As Boolean

Private Sub Form_DblClick()

    If Me.WindowState <> vbNormal Then Exit Sub
    Me.Width = Me.Height * 1.33
    
End Sub

Private Sub Form_Load()

    Call Form_DblClick
    
End Sub

Private Sub Form_Resize()

    Me.pictMap.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Call Init_ViewMapping
    
End Sub

Private Sub pictMap_DblClick()

    If Me.WindowState <> vbNormal Then Exit Sub
    Me.Width = Me.Height * 1.33
    
End Sub

Private Sub pictMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_blnMouseDown = True

End Sub

Private Sub pictMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    y = -y
    If m_blnMouseDown = True Then
        m_Window.xMin = x
        m_Window.xMax = x + Shape1.Width
        
        m_Window.yMin = y
        m_Window.yMax = y + Shape1.Height
        
        Call Init_ViewMapping
    End If
    
    Me.Caption = "x: " & Int(x) & "  y: " & Int(y)
    
End Sub

Private Sub pictMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_blnMouseDown = False
    
End Sub


