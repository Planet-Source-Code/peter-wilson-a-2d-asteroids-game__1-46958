VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BackColor       =   &H00404040&
   Caption         =   "Game Options"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   330
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0000
            Key             =   "Asteroid"
            Object.Tag             =   "Asteroid"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   9340
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Asteroids"
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   1380
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   930
         TabIndex        =   2
         Top             =   1080
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim itmX As ListItem
    
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid1", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid2", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid3", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid4", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid5", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid6", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid7", "Asteroid")
    Set itmX = Me.ListView1.ListItems.Add(, , "Asteroid8", "Asteroid")
    
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Call ReleaseCapture

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then ReleaseCapture

End Sub


