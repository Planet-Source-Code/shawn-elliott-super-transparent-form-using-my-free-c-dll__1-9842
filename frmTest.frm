VERSION 5.00
Object = "{0576E0EE-5C84-11D4-B234-0080C8F8E0E3}#4.0#0"; "TransRegion.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TransRegionCtrl.TransRegion TransRegion1 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   6006
      Picture         =   "frmTest.frx":0000
      Begin VB.CommandButton Command2 
         Caption         =   "R"
         Height          =   225
         Left            =   3600
         TabIndex        =   4
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "G"
         Height          =   225
         Left            =   3855
         TabIndex        =   3
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "B"
         Height          =   225
         Left            =   4110
         TabIndex        =   2
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   270
         Left            =   3645
         TabIndex        =   1
         Top             =   555
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans"
         Height          =   225
         Left            =   3765
         TabIndex        =   5
         Top             =   1050
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End

End Sub

Private Sub Command2_Click()
'We change the Mask colors so we can create multiple transparencies
'Very important.  If you enter Mask properties you must send 0,0,0 to the
'TransRegion Method

With TransRegion1
    .MaskRed = 255
    .MaskBlue = 0
    .MaskGreen = 0
End With

TransRegion1.TransRegion 0, 0, 0
End Sub

Private Sub Command3_Click()

With TransRegion1
    .MaskRed = 0
    .MaskBlue = 0
    .MaskGreen = 255
End With

TransRegion1.TransRegion 0, 0, 0
End Sub

Private Sub Command4_Click()
With TransRegion1
    .MaskRed = 0
    .MaskBlue = 255
    .MaskGreen = 0
End With

TransRegion1.TransRegion 0, 0, 0
End Sub

Private Sub Form_Load()
TransRegion1.TransRegion 0, 0, 0


End Sub

