VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DaBooda 2D Engine v1.0 Example"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   616
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox DPic 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   0
      ScaleHeight     =   7200
      ScaleWidth      =   7200
      TabIndex        =   1
      Top             =   2040
      Width           =   7200
   End
   Begin VB.PictureBox DaBoodaLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   0
      Picture         =   "Main.frx":0CCA
      ScaleHeight     =   2040
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Main.Show         'Show form
    DoEvents        'This just lets the form catch up
    MsgBox "Welcome to DaBooda2D!  The Arrow Keys move the map.  A and R add and remove sprites.  + and - change framerate.  Escape quits.  Enjoy!"
            
    Initialize      'This starts the example in motion
    
'Everything in the example is done through the example module
'I like setting up programs in modules.....just my preference
'You don't have to do this
    
End Sub
