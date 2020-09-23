VERSION 5.00
Begin VB.Form frmTestProgress 
   Caption         =   "Progress Bar Control Tester"
   ClientHeight    =   7455
   ClientLeft      =   5940
   ClientTop       =   4710
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStep 
      Caption         =   "&Step"
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "&Run"
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Top             =   7020
      Width           =   1155
   End
   Begin VB.Timer tmrUpd 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2040
      Top             =   7020
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Tag             =   "1"
      Top             =   0
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":1272
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":1290
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":1600
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar2 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Tag             =   "2"
      Top             =   360
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":1628
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":1646
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":19B2
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar3 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Tag             =   "3"
      Top             =   720
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":19DA
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":19F8
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":1D98
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar4 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1080
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":1DC0
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":1DDE
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":214E
      segments        =   -1
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar5 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Tag             =   "2"
      Top             =   1440
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":2176
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":2194
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":2500
      segments        =   -1
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar6 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Tag             =   "3"
      Top             =   1800
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":2528
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":2546
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":28E6
      segments        =   -1
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar7 
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Tag             =   "1"
      Top             =   6780
      Width           =   4275
      _extentx        =   7541
      _extenty        =   344
      picture         =   "frmTest.frx":290E
      backcolor       =   0
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":2CAC
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":3018
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar8 
      Height          =   2775
      Left            =   1680
      TabIndex        =   9
      Tag             =   "1"
      Top             =   3960
      Width           =   435
      _extentx        =   767
      _extenty        =   4895
      picture         =   "frmTest.frx":3040
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":305E
      barpicturemode  =   0
      showtext        =   -1
      font            =   "frmTest.frx":33D2
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar9 
      Height          =   2775
      Left            =   2640
      TabIndex        =   10
      Tag             =   "2"
      Top             =   3960
      Width           =   435
      _extentx        =   767
      _extenty        =   4895
      picture         =   "frmTest.frx":33FA
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":3418
      barpicturemode  =   0
      showtext        =   -1
      font            =   "frmTest.frx":3798
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar10 
      Height          =   2775
      Left            =   2160
      TabIndex        =   11
      Tag             =   "3"
      Top             =   3960
      Width           =   435
      _extentx        =   767
      _extenty        =   4895
      picture         =   "frmTest.frx":37C0
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":37DE
      barpicturemode  =   0
      showtext        =   -1
      font            =   "frmTest.frx":3B42
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar11 
      Height          =   2775
      Left            =   1200
      TabIndex        =   12
      Tag             =   "1"
      Top             =   3960
      Width           =   435
      _extentx        =   767
      _extenty        =   4895
      picture         =   "frmTest.frx":3B6A
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":3EB0
      barpicturemode  =   0
      showtext        =   -1
      font            =   "frmTest.frx":4214
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar12 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Tag             =   "1"
      Top             =   2160
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":423C
      backcolor       =   0
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":425A
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":45CA
      segments        =   -1
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar13 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Tag             =   "2"
      Top             =   2520
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":45F2
      backcolor       =   0
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":4610
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":497C
      segments        =   -1
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar14 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Tag             =   "3"
      Top             =   2880
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":49A4
      backcolor       =   0
      forecolor       =   0
      appearance      =   0
      barpicture      =   "frmTest.frx":49C2
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":4D62
      segments        =   -1
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar15 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Tag             =   "2"
      Top             =   3240
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":4D8A
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":4DA8
      barpicturemode  =   0
      showtext        =   -1
      font            =   "frmTest.frx":5078
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar16 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Tag             =   "4"
      Top             =   3600
      Width           =   4215
      _extentx        =   7435
      _extenty        =   556
      picture         =   "frmTest.frx":50A0
      backcolor       =   16777215
      forecolor       =   0
      appearance      =   0
      barcolor        =   16777215
      barpicture      =   "frmTest.frx":5406
      barpicturemode  =   0
      backpicturemode =   0
      showtext        =   -1
      font            =   "frmTest.frx":5776
   End
End
Attribute VB_Name = "frmTestProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You will love it!
'The Best Progress Bar even created, very nice and powerfull.
'Less than perfect. ;-)
'3D look, Gold, image backgrounds... and plus.
'Created by Steve McMahon  f rom VBAccelerator.
'With images and form implementation by me - José Luis Farías
'JoseloFarias[at]adinet.com.uy
'Desde Uruguay.
'Please, if you use in your own proyects, please sendme a program copy (source code if better)
Option Explicit
Private Sub cmdAnimate_Click()
   If tmrUpd.Enabled Then
      tmrUpd.Enabled = False
      cmdStep.Enabled = True
      cmdAnimate.Caption = "&Run"
   Else
      tmrUpd.Enabled = True
      cmdStep.Enabled = False
      cmdAnimate.Caption = "&Stop"
   End If
End Sub
Private Sub cmdStep_Click()
   tmrUpd_Timer
End Sub
Private Sub tmrUpd_Timer()
Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is vbalProgressBar Then
         With ctl
            .Value = .Value + .Tag
            If ctl.ShowText Then
               If ctl.Name = "vbalProgressBar4" Then
                  .Text = .Value & "% Completed"
               Else
                  .Text = CLng(.Percent) & "%"
               End If
            End If
            If .Value >= .Max Then
               .Tag = -1 * Abs(.Tag)
            ElseIf .Value < 1 Then
               .Tag = Abs(.Tag)
            End If
         End With
      End If
   Next
End Sub
