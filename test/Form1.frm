VERSION 5.00
Object = "*\A..\mm_CheckBox.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00FFDBBF&
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFDBBF&
      Caption         =   "Frame1"
      Height          =   1890
      Left            =   150
      TabIndex        =   20
      Top             =   4785
      Width           =   6870
      Begin VB.CommandButton Command3 
         Caption         =   "&Enabled true/false"
         Height          =   675
         Left            =   4635
         TabIndex        =   31
         Top             =   675
         Width           =   1680
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   0
         Left            =   795
         TabIndex        =   22
         Top             =   255
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFDBBF&
         Caption         =   "small"
         Height          =   450
         Left            =   105
         TabIndex        =   21
         Top             =   -120
         Width           =   765
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   2
         Left            =   3045
         TabIndex        =   23
         Top             =   1350
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   3
         Left            =   3045
         TabIndex        =   24
         Top             =   825
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   4
         Left            =   3045
         TabIndex        =   25
         Top             =   255
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   5
         Left            =   795
         TabIndex        =   26
         Top             =   825
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   6
         Left            =   795
         TabIndex        =   27
         Top             =   1350
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   7
         Left            =   1965
         TabIndex        =   28
         Top             =   255
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   8
         Left            =   1965
         TabIndex        =   29
         Top             =   825
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   26
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   9
         Left            =   1965
         TabIndex        =   30
         Top             =   1350
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Small           =   0   'False
         RoundedValue    =   26
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4605
      Left            =   3930
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4605
      ScaleWidth      =   7425
      TabIndex        =   4
      Top             =   15
      Width           =   7425
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   0
         Left            =   5025
         TabIndex        =   5
         Top             =   300
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   6
         Left            =   5025
         TabIndex        =   6
         Top             =   570
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   7
         Left            =   5025
         TabIndex        =   7
         Top             =   840
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   8
         Left            =   5025
         TabIndex        =   8
         Top             =   1110
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   9
         Left            =   5025
         TabIndex        =   9
         Top             =   1380
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   10
         Left            =   5025
         TabIndex        =   10
         Top             =   1650
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   11
         Left            =   5025
         TabIndex        =   11
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   12
         Left            =   5025
         TabIndex        =   12
         Top             =   2190
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   13
         Left            =   5025
         TabIndex        =   13
         Top             =   2460
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   14
         Left            =   5025
         TabIndex        =   14
         Top             =   2730
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   15
         Left            =   5025
         TabIndex        =   15
         Top             =   3000
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   16
         Left            =   5025
         TabIndex        =   16
         Top             =   3255
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   17
         Left            =   5025
         TabIndex        =   17
         Top             =   3540
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   18
         Left            =   5025
         TabIndex        =   18
         Top             =   3810
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox2 
         Height          =   225
         Index           =   19
         Left            =   5025
         TabIndex        =   19
         Top             =   4080
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         RoundedValue    =   10
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H00C0C0FF&
         Height          =   4440
         Left            =   4545
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "etat"
      Height          =   345
      Left            =   2760
      TabIndex        =   1
      Top             =   2010
      Width           =   720
   End
   Begin MM_Advanced_CheckBox_v1.mm_checkbox mm_checkbox1 
      Height          =   450
      Left            =   1500
      TabIndex        =   2
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   794
      Checked         =   -1  'True
      Small           =   0   'False
      RoundedValue    =   26
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4605
      Left            =   15
      Picture         =   "Form1.frx":6F8B2
      ScaleHeight     =   4605
      ScaleWidth      =   3930
      TabIndex        =   3
      Top             =   15
      Width           =   3930
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         FillColor       =   &H00C0C0FF&
         Height          =   960
         Left            =   1020
         Shape           =   4  'Rounded Rectangle
         Top             =   1620
         Width           =   2580
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFDBBF&
      Caption         =   "&check All"
      Height          =   315
      Left            =   8775
      TabIndex        =   0
      Top             =   4650
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

On Error Resume Next

For i = 0 To 20
    Me.mm_checkbox2(i).Checked = Check1.Value
Next i

End Sub

Private Sub Check2_Click()
On Error Resume Next
For i = 0 To 9
    Me.mm_checkbox3(i).Small = Check2.Value
    Me.mm_checkbox3(i).Enabled = True
Next i

End Sub

Private Sub Command1_Click()

MsgBox mm_checkbox1.Checked

End Sub

Private Sub Command3_Click()
On Error Resume Next
For i = 0 To 9
    mm_checkbox3(i).Enabled = Not mm_checkbox3(i).Enabled
Next i

End Sub


Private Sub Form_Load()

On Error Resume Next



For i = 0 To 20
    Randomize
    Me.mm_checkbox2(i).Checked = Int((2 * Rnd()) - 1)
Next i


For i = 0 To 9
    Me.mm_checkbox3(i).RoundedValue = 31
Next i

End Sub




