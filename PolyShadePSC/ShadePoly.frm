VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Polygon Shader"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   10320
   DrawStyle       =   5  'Transparent
   Icon            =   "ShadePoly.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Container Container7 
      Height          =   2865
      Left            =   210
      TabIndex        =   24
      Top             =   105
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   5054
      BackColor       =   14737632
      Caption         =   "Colors"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":0E42
      Begin VB.CommandButton cmdSwap 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   2715
         Picture         =   "ShadePoly.frx":0E5E
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   " Copy "
         Top             =   1050
         Width           =   210
      End
      Begin VB.PictureBox picRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   1305
         ScaleHeight     =   165
         ScaleWidth      =   1035
         TabIndex        =   41
         Top             =   1950
         Width           =   1035
      End
      Begin VB.PictureBox picRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   1290
         ScaleHeight     =   165
         ScaleWidth      =   1035
         TabIndex        =   40
         Top             =   1095
         Width           =   1035
      End
      Begin VB.PictureBox picRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   1290
         ScaleHeight     =   165
         ScaleWidth      =   1035
         TabIndex        =   39
         Top             =   255
         Width           =   1035
      End
      Begin VB.PictureBox picCul 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   135
         Picture         =   "ShadePoly.frx":0F30
         ScaleHeight     =   480
         ScaleWidth      =   2160
         TabIndex        =   37
         Top             =   2205
         Width           =   2190
      End
      Begin VB.PictureBox picCul 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   135
         Picture         =   "ShadePoly.frx":4572
         ScaleHeight     =   480
         ScaleWidth      =   2160
         TabIndex        =   32
         Top             =   1350
         Width           =   2190
      End
      Begin VB.CommandButton cmdSwap 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   2385
         Picture         =   "ShadePoly.frx":7BB4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   " Swap "
         Top             =   1050
         Width           =   270
      End
      Begin VB.PictureBox picCul 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   135
         Picture         =   "ShadePoly.frx":7C86
         ScaleHeight     =   480
         ScaleWidth      =   2160
         TabIndex        =   27
         Top             =   480
         Width           =   2190
      End
      Begin VB.Label LabCul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   2460
         TabIndex        =   36
         Top             =   2190
         Width           =   435
      End
      Begin VB.Label LabC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   765
         TabIndex        =   35
         Top             =   1980
         Width           =   450
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BACK"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   34
         Top             =   1965
         Width           =   570
      End
      Begin VB.Label LabCul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   2460
         TabIndex        =   33
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label LabC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   765
         TabIndex        =   30
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "END"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   1095
         Width           =   465
      End
      Begin VB.Label LabCul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   2445
         TabIndex        =   28
         Top             =   480
         Width           =   405
      End
      Begin VB.Label LabC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   0
         Left            =   750
         TabIndex        =   26
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "START"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   240
         Width           =   570
      End
   End
   Begin Project1.Container Container6 
      Height          =   1125
      Left            =   225
      TabIndex        =   20
      Top             =   5760
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1984
      BackColor       =   14737632
      BorderColorDark =   255
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":B2C8
      Begin VB.CheckBox chkPoint 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   150
         TabIndex        =   21
         Top             =   165
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Left-Click on points  -> Right-Click to end"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   465
         TabIndex        =   23
         Top             =   120
         Width           =   2445
      End
      Begin VB.Label LabMNP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "LabMNP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   750
         Width           =   2355
      End
   End
   Begin Project1.Container Container5 
      Height          =   600
      Left            =   1965
      TabIndex        =   17
      Top             =   5040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1058
      BackColor       =   14737632
      Caption         =   "Points X,Y"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":B2E4
      Begin VB.Label LabXY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabXY"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   315
         TabIndex        =   19
         Top             =   255
         Width           =   825
      End
      Begin VB.Label LabPNum 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   90
         TabIndex        =   18
         Top             =   270
         Width           =   210
      End
   End
   Begin Project1.Container Container4 
      Height          =   660
      Left            =   225
      TabIndex        =   14
      ToolTipText     =   " +ve Polys, -ve Stars "
      Top             =   4290
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1164
      BackColor       =   14737632
      Caption         =   "Regular +/- "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":B300
      Begin VB.ComboBox cboStarArm 
         Height          =   315
         Left            =   2055
         TabIndex        =   43
         Text            =   "0.5"
         Top             =   225
         Width           =   765
      End
      Begin VB.ComboBox cboRegular 
         Height          =   315
         ItemData        =   "ShadePoly.frx":B31C
         Left            =   540
         List            =   "ShadePoly.frx":B31E
         TabIndex        =   16
         Text            =   "3"
         Top             =   240
         Width           =   630
      End
      Begin VB.CheckBox chkRegular 
         Caption         =   "Check1"
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   285
         Width           =   225
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Star   Arm"
         Height          =   405
         Left            =   1515
         TabIndex        =   44
         Top             =   150
         Width           =   390
      End
   End
   Begin Project1.Container Container3 
      Height          =   600
      Left            =   1800
      TabIndex        =   11
      Top             =   3645
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   1058
      BackColor       =   14737632
      Caption         =   "'Center' point"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":B320
      Begin VB.OptionButton optCenPt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Man"
         Height          =   270
         Index           =   1
         Left            =   750
         TabIndex        =   13
         Top             =   225
         Width           =   675
      End
      Begin VB.OptionButton optCenPt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto"
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   225
         Width           =   750
      End
   End
   Begin Project1.Container Container2 
      Height          =   600
      Left            =   195
      TabIndex        =   8
      Top             =   3645
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   1058
      BackColor       =   14737632
      Caption         =   "Gradient"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":B33C
      Begin VB.OptionButton optGrad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hue"
         Height          =   210
         Index           =   1
         Left            =   810
         TabIndex        =   10
         Top             =   270
         Width           =   615
      End
      Begin VB.OptionButton optGrad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Linear"
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   9
         Top             =   270
         Width           =   810
      End
   End
   Begin Project1.Container Container1 
      Height          =   585
      Left            =   210
      TabIndex        =   4
      Top             =   2970
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   1032
      BackColor       =   14737632
      BorderColorLight=   14737632
      Caption         =   "Track"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "ShadePoly.frx":B358
      Begin VB.CommandButton cmdRUR 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete"
         Height          =   285
         Index           =   2
         Left            =   2115
         TabIndex        =   38
         Top             =   210
         Width           =   750
      End
      Begin VB.CommandButton cmdRUR 
         BackColor       =   &H00E0E0E0&
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   " Forward "
         Top             =   210
         Width           =   555
      End
      Begin VB.CommandButton cmdRUR 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   " Back "
         Top             =   210
         Width           =   555
      End
      Begin VB.Label LabTNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1515
         TabIndex        =   7
         Top             =   225
         Width           =   465
      End
   End
   Begin VB.CheckBox chkHairs 
      BackColor       =   &H00E0E0E0&
      Caption         =   " + Hairs"
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   5205
      Width           =   945
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   6885
      Left            =   3855
      ScaleHeight     =   455
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   0
      Top             =   300
      Width           =   6000
      Begin VB.Shape shpPt 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         Height          =   90
         Index           =   0
         Left            =   1560
         Shape           =   3  'Circle
         Top             =   255
         Width           =   105
      End
      Begin VB.Line LineV 
         BorderColor     =   &H80000009&
         BorderStyle     =   3  'Dot
         DrawMode        =   7  'Invert
         X1              =   54
         X2              =   54
         Y1              =   40
         Y2              =   81
      End
      Begin VB.Line LineH 
         BorderColor     =   &H80000009&
         BorderStyle     =   3  'Dot
         DrawMode        =   7  'Invert
         X1              =   20
         X2              =   81
         Y1              =   28
         Y2              =   28
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      Height          =   7140
      Index           =   1
      Left            =   45
      Top             =   60
      Width           =   3420
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000012&
      BorderWidth     =   2
      Height          =   7140
      Index           =   0
      Left            =   60
      Top             =   75
      Width           =   3420
   End
   Begin VB.Line Line6 
      X1              =   352
      X2              =   357
      Y1              =   14
      Y2              =   9
   End
   Begin VB.Line Line5 
      X1              =   358
      X2              =   349
      Y1              =   9
      Y2              =   4
   End
   Begin VB.Line Line4 
      X1              =   252
      X2              =   247
      Y1              =   78
      Y2              =   87
   End
   Begin VB.Line Line3 
      X1              =   246
      X2              =   241
      Y1              =   86
      Y2              =   77
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "X"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   6435
      TabIndex        =   2
      Top             =   30
      Width           =   150
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Y"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   3645
      TabIndex        =   1
      Top             =   1740
      Width           =   150
   End
   Begin VB.Line Line2 
      X1              =   247
      X2              =   247
      Y1              =   9
      Y2              =   113
   End
   Begin VB.Line Line1 
      X1              =   248
      X2              =   387
      Y1              =   9
      Y2              =   9
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuOpenSave 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuOpenSave 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuOpenSave 
         Caption         =   "&Open ply file"
         Index           =   2
      End
      Begin VB.Menu mnuOpenSave 
         Caption         =   "&Save ply file"
         Index           =   3
      End
      Begin VB.Menu mnuOpenSave 
         Caption         =   "Save &BMP"
         Index           =   4
      End
      Begin VB.Menu mnuOpenSave 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuOpenSave 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Polygon Shader by Robert Rayment  Dec 2004


Option Explicit
Option Base 1

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private ORGFormWidth As Long
Private ORGFormHeight As Long
Private PICW As Long, PICH As Long

Private NumPoints As Long     ' in current polygon
Private SColor As Long, EColor As Long    ' Start and End colors
Private BColor As Long        ' Back color
Private PolygonNumber As Long, MaxNumPolygons As Long
Private xp() As Single, yp() As Single    ' Polygon points
Private GradType As Long  ' 0 Linear, 1 Hue
Private aPlot As Boolean   ' To prevent click-thru

Private AutoManual As Long ' 0 Automatic, 1 Manual center point
Private aManualCenter As Boolean ' Flag to calc or use center shade point

Private aRegular As Boolean   ' Regular polygons
Private RegNumber As Long     ' Number of regular polygon sides (3-16)
Private zArmFrac As Single

'Saving
Private xsv() As Single, ysv() As Single    ' Saved polygons
Private xc() As Single, yc() As Single      ' Saved center point X,Y
Private SVNumPoints() As Long               ' Num points in polygon
Private SaveSC() As Long, SaveEC() As Long  ' Saved colors
Private LinHue() As Long         ' 0 Linear, 1 Hue gradient

Private STX As Long, STY As Long    ' Twips/pixel

Private PathSpec$, CurrentPath$, FileSpec$

Private Const MAXNUMPTS As Long = 32

Dim CommonDialog1 As OSDialog


Private Sub Form_Initialize()
   InitCommonControls
End Sub

Private Sub Form_Load()
Dim k As Long
   ORGFormWidth = Form1.Width
   ORGFormHeight = Form1.Height
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   PICW = 400
   PICH = 400
   With PIC
      .AutoRedraw = True
      .ScaleMode = vbPixels
      .Width = PICW
      .Height = PICH
   End With
   
   LabMNP = "Max num of points = +/-" & Str$(MAXNUMPTS \ 2)
   
   NumPoints = 0
   ReDim xp(MAXNUMPTS), yp(MAXNUMPTS)
   PolygonNumber = 0
   MaxNumPolygons = 0
   ReDim xsv(MAXNUMPTS, 1), ysv(MAXNUMPTS, 1)
   ReDim xc(1), yc(1)
   ReDim SVNumPoints(1)
   ReDim SaveSC(1), SaveEC(1)
   ReDim LinHue(1)
   ReDim xd(2), yd(2)
   SColor = 255
   EColor = 0
   BColor = PIC.BackColor
   LabCul(0).BackColor = SColor
   LabCul(1).BackColor = EColor
   
   chkHairs.Value = False
   LineH.Visible = False
   LineV.Visible = False
   LabXY = ""
   LabPNum = "0"
   ' Point markers
   For k = 1 To MAXNUMPTS - 1
      Load shpPt(k)
      shpPt(k).Visible = False
   Next k
   shpPt(0).Visible = False
   
   optGrad(0) = True
   GradType = 0   ' Linear
   chkPoint.Value = Checked
   aPlot = True
   AutoManual = 0
   aManualCenter = False
   
   aRegular = False
   chkRegular.Value = Unchecked
   For k = -16 To -3
      cboRegular.AddItem (Str$(k))
   Next k
   For k = 3 To 16
      cboRegular.AddItem (Str$(k))
   Next k
   RegNumber = 3
   
   For k = 1 To 20
      'v$ = Str$(k / 8)
      cboStarArm.AddItem Str$(k / 20)
   Next k
   cboStarArm.Text = "0.5"
   zArmFrac = 0.5
   
   optCenPt(0).Value = True
   cmdRUR(0).Enabled = False  ' Undo
   cmdRUR(1).Enabled = False  ' Redo
   cmdRUR(2).Enabled = False  ' Delete
   
   'cmdSwap.Caption = Chr$(128)
End Sub

Private Sub CenterShadePolygon(X As Single, Y As Single)
' Input
' aManualCenter = False: returns shade point at X=x3, Y=y3
' aManualCenter = True:  forces shade point to be x3=X, y3=Y

' xp(1 to NumPoints),yp(1 to NumPoints) ' Polygon coords
' Point NumPoints is the last and is set to the start
' color.

' aRegular And RegNumber > 0 gives regular polygons
' aRegular And RegNumber < 0 gives regular stars

' Uses global Private: SColor, EColor & GradType

ReDim xa(NumPoints) As Single, ya(NumPoints) As Single

Dim X3 As Single, Y3 As Single
Dim k As Long, j As Long, n As Long
Dim kup As Long
   
   If Not aManualCenter Then
      For k = 1 To NumPoints
         j = k + 1
         If j > NumPoints Then j = 1
         xa(k) = (xp(k) + xp(j)) / 2
         ya(k) = (yp(k) + yp(j)) / 2
      Next k
      For n = 1 To MAXNUMPTS
         For k = 1 To NumPoints
            j = k + 1
            If j > NumPoints Then j = 1
            xa(k) = (xa(k) + xa(j)) / 2
            ya(k) = (ya(k) + ya(j)) / 2
         Next k
      Next n
      X3 = xa(k - 1)
      Y3 = ya(k - 1)
      X = X3
      Y = Y3
   Else
      X3 = X
      Y3 = Y
   End If
   
   
   
   If aRegular And RegNumber < 0 Then
      ' Develop star points between inputted points
      ReDim xa(2 * NumPoints) As Single, ya(2 * NumPoints) As Single
      For k = 1 To NumPoints
         xa(2 * k - 1) = xp(k)
         ya(2 * k - 1) = yp(k)
      Next k
      NumPoints = 2 * NumPoints
      For k = 2 To NumPoints Step 2
         kup = k + 1
         If kup > NumPoints Then kup = 1
         xa(k) = (xa(k - 1) + xa(kup)) / 2
         ya(k) = (ya(k - 1) + ya(kup)) / 2
      Next k
      For k = 2 To NumPoints Step 2
         xa(k) = X3 + Abs(xa(k) - X3) * zArmFrac * Sgn(xa(k) - X3)
         ya(k) = Y3 + Abs(ya(k) - Y3) * zArmFrac * Sgn(ya(k) - Y3)
      Next k
      For k = 1 To NumPoints
         xp(k) = xa(k)
         yp(k) = ya(k)
      Next k
   End If
   
   
   If aManualCenter Then
   If NumPoints = 3 Then   ' Triangle simple shading
   If X3 = xp(3) Then
   If Y3 = yp(3) Then
      CenterShadeTriangle xp(1), yp(1), _
                     xp(2), yp(2), _
                     X3, Y3
      Exit Sub
   End If
   End If
   End If
   End If
   
   For k = 1 To NumPoints
      j = k + 1
      If j > NumPoints Then j = 1
      CenterShadeTriangle xp(k), yp(k), _
                     xp(j), yp(j), _
                     X3, Y3
   Next k
   LabXY = ""
   LabPNum = "0"
   For k = 0 To MAXNUMPTS - 1
      shpPt(k).Visible = False
   Next k

End Sub

Private Sub CenterShadeTriangle(xi1 As Single, yi1 As Single, _
                           xi2 As Single, yi2 As Single, _
                           xi3 As Single, yi3 As Single)

' Uses global Private: SColor, EColor & GradType

Dim zphi As Single  ' Angles to calc steps
Dim zSlantDis1 As Single   ' Dis 3 -> 1
Dim zSlantDis2 As Single   ' Dis 3 -> 2
Dim zMaxSlant As Single
Dim xd(2) As Single, yd(2) As Single    ' x & y steps along lines 3->1 & 3->2
Dim NSteps As Long
' Colors
Dim bR As Byte, bG As Byte, bB As Byte
Dim zSR As Single, zSG As Single, zSB As Single
Dim zER As Single, zEG As Single, zEB As Single
Dim zR As Single, zG As Single, zB As Single
Dim zdR As Single, zdG As Single, zdB As Single
'
Dim k As Long
Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
      
' Hue, Saturation, Luminance  Start & End
Dim zHS As Single, zSS As Single, zLS As Single
Dim zHE As Single, zSE As Single, zLE As Single
Dim zHueID As Single
Dim zHue As Single, zSat As Single, zLum As Single
Dim zHueIncDec As Single, zSatIncDec As Single, zLumIncDec As Single
   
   zSlantDis1 = Sqr((xi1 - xi3) ^ 2 + (yi1 - yi3) ^ 2)
   If zSlantDis1 = 0 Then zSlantDis1 = 1
   zSlantDis2 = Sqr((xi2 - xi3) ^ 2 + (yi2 - yi3) ^ 2)
   If zSlantDis2 = 0 Then zSlantDis2 = 1
   zMaxSlant = zSlantDis1
   If zSlantDis1 > zSlantDis2 Then
      zMaxSlant = zSlantDis1
      zphi = 2 * pi# - zATan2(yi1 - yi3, xi1 - xi3)
      yd(1) = Sin(zphi)
      xd(1) = Cos(zphi)
      zphi = 2 * pi - zATan2(yi2 - yi3, xi2 - xi3)
      yd(2) = Sin(zphi) * zSlantDis2 / zSlantDis1
      xd(2) = Cos(zphi) * zSlantDis2 / zSlantDis1
   Else
      zMaxSlant = zSlantDis2
      zphi = 2 * pi - zATan2(yi2 - yi3, xi2 - xi3)
      yd(2) = Sin(zphi)
      xd(2) = Cos(zphi)
      zphi = 2 * pi# - zATan2(yi1 - yi3, xi1 - xi3)
      yd(1) = Sin(zphi) * zSlantDis1 / zSlantDis2
      xd(1) = Cos(zphi) * zSlantDis1 / zSlantDis2
   End If
   NSteps = CInt(zMaxSlant)
   
   PIC.DrawWidth = 2
   ' Start color  ' SColor
   LngToRGB SColor, bR, bG, bB
   zSR = bR
   zSG = bG
   zSB = bB
   zR = zSR
   zG = zSG
   zB = zSB
   ' End color ' EColor
   LngToRGB EColor, bR, bG, bB
   zER = bR
   zEG = bG
   zEB = bB
   
   If GradType = 0 Then    ' Linear
      ' Linear RGB increments
      zdR = (zER - zSR) / NSteps
      zdG = (zEG - zSG) / NSteps
      zdB = (zEB - zSB) / NSteps
      For k = 1 To NSteps
         x1 = xi3 + (k - 1) * xd(1)
         y1 = yi3 + (k - 1) * yd(1)
         x2 = xi3 + (k - 1) * xd(2)
         y2 = yi3 + (k - 1) * yd(2)
         PIC.Line (x1, y1)-(x2, y2), RGB(zR, zG, zB)
         zR = zR + zdR
         zG = zG + zdG
         zB = zB + zdB

         If zR > 255 Then zR = 255
         If zG > 255 Then zG = 255
         If zB > 255 Then zB = 255
         If zR < 0 Then zR = 0
         If zG < 0 Then zG = 0
         If zB < 0 Then zB = 0
'OR
'         If zR > 255 Then zR = zR - 255
'         If zG > 255 Then zG = zG - 255
'         If zB > 255 Then zB = zB - 255
'         If zR < 0 Then zR = zR + 255
'         If zG < 0 Then zG = zG + 255
'         If zB < 0 Then zB = zB + 255
      Next k
   
   Else ' Index = 1 Hue gradient
   
      
      RGB2HSL zSR, zSG, zSB, zHS, zSS, zLS
      RGB2HSL zER, zEG, zEB, zHE, zSE, zLE
   
      zHueID = zHE - zHS
      If Abs(zHueID) > 127.5 Then
         zHueID = (255 - Abs(zHueID)) * -Sgn(zHueID)
      End If
      zHueIncDec = zHueID / NSteps
      zSatIncDec = (zSE - zSS) / NSteps
      zLumIncDec = (zLE - zLS) / NSteps
      
      zHue = zHS
      zSat = zSS
      zLum = zLS
   
      For k = 1 To NSteps
         HSL2RGB zHue, zSat, zLum, zR, zG, zB
         x1 = xi3 + (k - 1) * xd(1)
         y1 = yi3 + (k - 1) * yd(1)
         x2 = xi3 + (k - 1) * xd(2)
         y2 = yi3 + (k - 1) * yd(2)
         PIC.Line (x1, y1)-(x2, y2), RGB(zR, zG, zB)
         
         zHue = zHue + zHueIncDec
         zSat = zSat + zSatIncDec
         zLum = zLum + zLumIncDec
      
         Select Case zHue
         Case Is < 0
            zHue = zHue + 255
         Case Is > 255
            zHue = zHue - 255
         End Select
      Next k
   End If
   
   PIC.DrawWidth = 1
End Sub

Private Sub cmdRUR_Click(Index As Integer)
Dim k As Long, j As Long
   
   If NumPoints <> 0 Then Exit Sub
   
   For k = 0 To MAXNUMPTS - 1
      shpPt(k).Visible = False
   Next k
   
   Select Case Index
   Case 0      ' Track back
      If PolygonNumber > 1 Then
         PolygonNumber = PolygonNumber - 1
         NumPoints = SVNumPoints(PolygonNumber)
         For k = 1 To NumPoints
            xp(k) = xsv(k, PolygonNumber)
            yp(k) = ysv(k, PolygonNumber)
            shpPt(k - 1).Left = xp(k) - 2
            shpPt(k - 1).Top = yp(k) - 2
            shpPt(k - 1).Visible = True
         Next k
         NumPoints = 0
         If PolygonNumber = 1 Then
            cmdRUR(0).Enabled = False  ' Track back
         End If
         cmdRUR(1).Enabled = True  ' Track fwrd
      End If
   Case 1      ' Track forward
      If PolygonNumber < MaxNumPolygons Then
         PolygonNumber = PolygonNumber + 1
         NumPoints = SVNumPoints(PolygonNumber)
         For k = 1 To NumPoints
            xp(k) = xsv(k, PolygonNumber)
            yp(k) = ysv(k, PolygonNumber)
            shpPt(k - 1).Left = xp(k) - 2
            shpPt(k - 1).Top = yp(k) - 2
            shpPt(k - 1).Visible = True
         Next k
         NumPoints = 0
         If PolygonNumber = MaxNumPolygons Then
            cmdRUR(1).Enabled = False  ' Track fwrd
         End If
         cmdRUR(0).Enabled = True  ' Track back
      End If
   Case 2      ' Delete
      If PolygonNumber = MaxNumPolygons Then
         
         PolygonNumber = PolygonNumber - 1
         MaxNumPolygons = PolygonNumber
         If PolygonNumber > 0 Then
            ReDim Preserve xsv(MAXNUMPTS, MaxNumPolygons), ysv(MAXNUMPTS, MaxNumPolygons)
            ReDim Preserve xc(MaxNumPolygons), yc(MaxNumPolygons)
            ReDim Preserve SVNumPoints(MaxNumPolygons)
            ReDim Preserve LinHue(MaxNumPolygons)
            ReDim Preserve SaveSC(MaxNumPolygons), SaveEC(MaxNumPolygons)
            PIC.Cls
            DrawPolygons MaxNumPolygons
            NumPoints = SVNumPoints(PolygonNumber)
            For k = 1 To MAXNUMPTS
               xp(k) = xsv(k, PolygonNumber)
               yp(k) = ysv(k, PolygonNumber)
               shpPt(k - 1).Left = xp(k) - 2
               shpPt(k - 1).Top = yp(k) - 2
               shpPt(k - 1).Visible = True
            Next k
            NumPoints = 0
            cmdRUR(0).Enabled = True   ' Track back
            cmdRUR(1).Enabled = False  ' Track fwrd
            cmdRUR(2).Enabled = True   ' Delete
         Else  ' PolygonNumber = 0
            ReDim xsv(MAXNUMPTS, 1), ysv(MAXNUMPTS, 1)
            ReDim xc(1), yc(1)
            ReDim SVNumPoints(1)
            ReDim LinHue(1)
            ReDim SaveSC(1), SaveEC(1)
            cmdRUR(0).Enabled = False  ' Undo
            cmdRUR(1).Enabled = False  ' Redo
            cmdRUR(2).Enabled = False  ' Delete
            PIC.Cls
            NumPoints = 0
            aManualCenter = False
         End If
      
      Else  ' PolygonNumber < MaxNumPolygons
         
         If PolygonNumber > 0 Then
            If MaxNumPolygons = 1 Then
               MaxNumPolygons = 0
               PolygonNumber = 0
               ReDim xsv(MAXNUMPTS, 1), ysv(MAXNUMPTS, 1)
               ReDim xc(1), yc(1)
               ReDim SVNumPoints(1)
               ReDim LinHue(1)
               ReDim SaveSC(1), SaveEC(1)
               cmdRUR(0).Enabled = False  ' Undo
               cmdRUR(1).Enabled = False  ' Redo
               cmdRUR(2).Enabled = False  ' Delete
               PIC.Cls
               NumPoints = 0
               aManualCenter = False
            Else  ' MaxNumPolygons > 1
               ' Overwrite PolygonNumber with those above
               For k = PolygonNumber + 1 To MaxNumPolygons
                  For j = 1 To MAXNUMPTS
                     xsv(j, k - 1) = xsv(j, k)
                     ysv(j, k - 1) = ysv(j, k)
                  Next j
                  xc(k - 1) = xc(k)
                  yc(k - 1) = yc(k)
                  SVNumPoints(k - 1) = SVNumPoints(k)
                  LinHue(k - 1) = LinHue(k)
                  SaveSC(k - 1) = SaveSC(k)
                  SaveEC(k - 1) = SaveEC(k)
               Next k
               ' Clear top polygon
               MaxNumPolygons = MaxNumPolygons - 1
               PolygonNumber = MaxNumPolygons
               ReDim Preserve xsv(MAXNUMPTS, MaxNumPolygons), ysv(MAXNUMPTS, MaxNumPolygons)
               ReDim Preserve xc(MaxNumPolygons), yc(MaxNumPolygons)
               ReDim Preserve SaveSC(MaxNumPolygons), SaveEC(MaxNumPolygons)
               ReDim Preserve SVNumPoints(MaxNumPolygons)
               ReDim Preserve LinHue(MaxNumPolygons)
               PIC.Cls
               DrawPolygons PolygonNumber
               NumPoints = SVNumPoints(PolygonNumber)
               For k = 1 To NumPoints
                  xp(k) = xsv(k, PolygonNumber)
                  yp(k) = ysv(k, PolygonNumber)
                  shpPt(k - 1).Left = xp(k) - 2
                  shpPt(k - 1).Top = yp(k) - 2
                  shpPt(k - 1).Visible = True
               Next k
               NumPoints = 0
               cmdRUR(0).Enabled = True   ' Track back
               cmdRUR(1).Enabled = False  ' Track fwrd
               cmdRUR(2).Enabled = True   ' Delete
            End If
         End If
      End If
   End Select
   LabTNum = Str$(PolygonNumber)
End Sub

Private Sub DrawPolygons(n As Long)
' Called from:-
' cmdRUR_Click       ' Undo/Redo
' LoadPlyFile
' picCul_MouseUp 2   ' BackColor
' Form_Resize

' n = PolygonNumber

Dim svSColor As Long, svEColor As Long
Dim svGradType As Long
Dim svRegNumber As Long
Dim svaManualCenter As Boolean
Dim k As Long, j As Long
   svSColor = SColor
   svEColor = EColor
   svGradType = GradType
   svaManualCenter = aManualCenter
   aManualCenter = True ' Force use of xc(), yc() centers
   svRegNumber = RegNumber
   RegNumber = Abs(RegNumber)
   For k = 1 To n
      NumPoints = SVNumPoints(k)
      For j = 1 To NumPoints
         xp(j) = xsv(j, k)
         yp(j) = ysv(j, k)
      Next j
      SColor = SaveSC(k)
      EColor = SaveEC(k)
      GradType = LinHue(k)
      CenterShadePolygon xc(k), yc(k)
   Next k
   SColor = svSColor
   EColor = svEColor
   GradType = svGradType
   aManualCenter = svaManualCenter
   RegNumber = svRegNumber
End Sub

Private Sub mnuOpenSave_Click(Index As Integer)
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim k As Long
   If NumPoints <> 0 Then Exit Sub
   ' To prevent click thru
   chkPoint.Value = Unchecked
   aPlot = False
   For k = 0 To MAXNUMPTS - 1
      shpPt(k).Visible = False
   Next k
   
   Select Case Index
   Case 0      ' New
      PIC.Cls
      PIC.BackColor = BColor
      NumPoints = 0
      ReDim xp(MAXNUMPTS), yp(MAXNUMPTS)
      PolygonNumber = 0
      MaxNumPolygons = 0
      ReDim xsv(MAXNUMPTS, 1), ysv(MAXNUMPTS, 1)
      ReDim xc(1), yc(1)
      cmdRUR(0).Enabled = False  ' Undo
      cmdRUR(1).Enabled = False  ' Redo
   Case 1      ' Break
   Case 2      ' Open ply file
      Set CommonDialog1 = New OSDialog
      Title$ = "Open triangles file"
      Filt$ = "Open ply|*.ply"
      InDir$ = CurrentPath$ 'Pathspec$
      FileSpec$ = ""
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FileSpec$) > 0 Then
         LoadPlyFile    'From FileSpec$
      End If
   Case 3      ' Save ply file
      If MaxNumPolygons = 0 Then
         MsgBox "No polygons to save", vbInformation, "PolyShader"
         Exit Sub
      End If
      Set CommonDialog1 = New OSDialog
      Title$ = "Save polygons file"
      Filt$ = "Save ply|*.ply"
      InDir$ = CurrentPath$ 'Pathspec$
      FileSpec$ = ""
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FileSpec$) > 0 Then
         FixExtension FileSpec$, ".ply"
         SavePlyFile   'To FileSpec$
      End If
   Case 4      ' Save BMP
      If MaxNumPolygons = 0 Then
         MsgBox "No polygons plotted", vbInformation, "TriShader"
         chkPoint.Value = Checked
         Exit Sub
      End If
      Set CommonDialog1 = New OSDialog
      Title$ = "Save As BMP"
      Filt$ = "Save bmp|*.bmp"
      InDir$ = CurrentPath$ 'Pathspec$
      FileSpec$ = ""
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FileSpec$) > 0 Then
         FixExtension FileSpec$, ".bmp"
         SavePicture PIC.Image, FileSpec$
      End If
   Case 5   ' Break
   Case 6   ' Exit
      Form_Unload 0
   End Select
   aManualCenter = False
End Sub

Private Sub SavePlyFile()
Dim FF As Long
Dim k As Long, j As Long
' xsv(MAXNUMPTS, MaxNumPolygons), ysv(MAXNUMPTS, MaxNumPolygons)
' SaveSC(MaxNumPolygons), SaveEC(MaxNumPolygons)
' LinHue(MaxNumPolygons)
' To FileSpec$
   FF = FreeFile
   Open FileSpec$ For Output As #FF
   Print #FF, FileSpec$, Date$
   Print #FF, MaxNumPolygons
   Print #FF, BColor
   For k = 1 To MaxNumPolygons
      Print #FF, SVNumPoints(k)
      Print #FF, LinHue(k)
      For j = 1 To SVNumPoints(k)
         Print #FF, xsv(j, k); " "; ysv(j, k)
      Next j
      Print #FF, SaveSC(k); " "; SaveEC(k)
      Print #FF, xc(k); " "; yc(k)
   Next k
   Close #FF
   NumPoints = 0
End Sub

Private Sub LoadPlyFile()
Dim FF As Long
Dim k As Long, j As Long
Dim A$

On Error GoTo LoadError
' xsv(MAXNUMPTS, MaxNumPolygons), ysv(MAXNUMPTS, MaxNumPolygons)
' SaveSC(MaxNumPolygons), SaveEC(MaxNumPolygons)
' LinHue(MaxNumPolygons)
' From FileSpec$
   FF = FreeFile
   Open FileSpec$ For Input As #FF
   Line Input #FF, A$
   Input #FF, MaxNumPolygons
   Input #FF, BColor
   PolygonNumber = MaxNumPolygons
   ReDim xsv(MAXNUMPTS, MaxNumPolygons), ysv(MAXNUMPTS, MaxNumPolygons)
   ReDim xc(MaxNumPolygons), yc(MaxNumPolygons)
   ReDim SaveSC(MaxNumPolygons), SaveEC(MaxNumPolygons)
   ReDim LinHue(MaxNumPolygons)
   ReDim SVNumPoints(MaxNumPolygons)
   For k = 1 To MaxNumPolygons
      Input #FF, SVNumPoints(k)
      Input #FF, LinHue(k)
      For j = 1 To SVNumPoints(k)
         Input #FF, xsv(j, k), ysv(j, k)
      Next j
      Input #FF, SaveSC(k), SaveEC(k)
      Input #FF, xc(k), yc(k)
   Next k
   Close #FF
   On Error GoTo 0
   
   PIC.Cls
   PIC.BackColor = BColor
   
   DrawPolygons MaxNumPolygons
   
   optCenPt(0).Value = True
   PolygonNumber = MaxNumPolygons
   
   NumPoints = SVNumPoints(PolygonNumber)
   For k = 1 To NumPoints
      xp(k) = xsv(k, PolygonNumber)
      yp(k) = ysv(k, PolygonNumber)
      shpPt(k - 1).Left = xp(k) - 2
      shpPt(k - 1).Top = yp(k) - 2
      shpPt(k - 1).Visible = True
   Next k
   NumPoints = 0
   cmdRUR(0).Enabled = True   ' Undo
   cmdRUR(1).Enabled = False  ' Track fwrd
   cmdRUR(2).Enabled = True  ' Delete
   LabTNum = Str$(PolygonNumber)
   Exit Sub
'==========
LoadError:
   Close
   On Error GoTo 0
   MsgBox "Load error", vbCritical, "Tri-shader"
   mnuOpenSave_Click 0
End Sub



Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim svaManualCenter As Boolean
Dim k As Long
Dim zL As Single
Dim zAng As Single, zRad As Double, zAlpha As Double
Dim xcen As Single, ycen As Single
Dim zcos As Double, zsin As Double
   
   If Not aPlot Then Exit Sub
   
   If aRegular Then
      
'----- REGULAR POLYS -------------------------------------------------------------
      If aManualCenter Then
         CenterShadePolygon X, Y    ' Shade current poly - @ X,Y
         StorePolygon X, Y          ' Bump Number of polygons and store
         aManualCenter = False
         Screen.MousePointer = vbDefault
      Else
         NumPoints = NumPoints + 1
         If NumPoints = 2 Then   ' calc the rest
            xp(NumPoints) = X
            yp(NumPoints) = Y
            shpPt(NumPoints - 1).Left = X - 2
            shpPt(NumPoints - 1).Top = Y - 2
            shpPt(NumPoints - 1).Visible = True
            ' 2 points defining length & angle of arm
            ' & RegNumber setting the number of arms 3->16
            ' Input length zL, radius = zR :-
            ' zL^2 = (x2-x1)^2 + (y2-y1)^2  AND
            ' zL^2 = 2*zR^2 - 2*zR^2 * cos(ang)   ie cos rule
            zL = Sqr((yp(2) - yp(1)) ^ 2 + (xp(2) - xp(1)) ^ 2)
            zAng = 2 * pi# - zATan2(yp(2) - yp(1), xp(2) - xp(1))
            zRad = 2 * pi# / Abs(RegNumber)
            zcos = Cos(zRad)
            zsin = Sin(zRad)
            'zR = zL/Sqr(2 * (1 - zcos))
            zAlpha = 0.5 * (pi# - zRad) - zAng
            'Apply cos rule
            xcen = xp(2) - zL * Cos(zAlpha) / Sqr(2 * (1 - zcos))
            ycen = yp(2) + zL * Sin(zAlpha) / Sqr(2 * (1 - zcos))
            
            For k = 3 To Abs(RegNumber)
               ' Rotate from previous point
               xp(k) = xcen + (xp(k - 1) - xcen) * zcos - (yp(k - 1) - ycen) * zsin
               yp(k) = ycen + (xp(k - 1) - xcen) * zsin + (yp(k - 1) - ycen) * zcos
               shpPt(k - 1).Left = xp(k) - 2    ' Show points
               shpPt(k - 1).Top = yp(k) - 2
               shpPt(k - 1).Visible = True
            Next k
            NumPoints = Abs(RegNumber)
            
            If AutoManual = 1 Then
               aManualCenter = True ' Next point will be manual center
               Screen.MousePointer = vbCrosshair
               Exit Sub
            End If
            
            aManualCenter = True
            CenterShadePolygon xcen, ycen    ' Shade current poly - Automatic
            StorePolygon xcen, ycen          ' Bump Number of polygons and store
            aManualCenter = False
         Else  ' First point on regular poly
            xp(NumPoints) = X
            yp(NumPoints) = Y
            shpPt(NumPoints - 1).Left = X - 2
            shpPt(NumPoints - 1).Top = Y - 2
            shpPt(NumPoints - 1).Visible = True
         End If
      End If
      Exit Sub    ' From regular poly
   End If
   
'----- NON-REGULAR POLYS ----------------------------------------------------------
   If aManualCenter Then
         CenterShadePolygon X, Y    ' Shade current poly - @ X,Y
         StorePolygon X, Y          ' Bump Number of polygons and store
         aManualCenter = False
         Screen.MousePointer = vbDefault
   Else
      NumPoints = NumPoints + 1
      If NumPoints = MAXNUMPTS Or Button = vbRightButton Then
         xp(NumPoints) = X
         yp(NumPoints) = Y
         ' Show points
         shpPt(NumPoints - 1).Left = X - 2
         shpPt(NumPoints - 1).Top = Y - 2
         shpPt(NumPoints - 1).Visible = True
         
         If AutoManual = 1 Then
            aManualCenter = True ' Next point will be manual center
            Screen.MousePointer = vbCrosshair
            Exit Sub
         End If
         
         CenterShadePolygon X, Y    ' Shade current poly - @ X,Y
         StorePolygon X, Y          ' Bump Number of polygons and store
      Else  'If NumPoints < MAXNUMPTS Then
         xp(NumPoints) = X
         yp(NumPoints) = Y
         shpPt(NumPoints - 1).Left = X - 2
         shpPt(NumPoints - 1).Top = Y - 2
         shpPt(NumPoints - 1).Visible = True
      End If
   End If
End Sub

Private Sub StorePolygon(X As Single, Y As Single)
Dim k As Long
   MaxNumPolygons = MaxNumPolygons + 1
   PolygonNumber = MaxNumPolygons
   ReDim Preserve xsv(MAXNUMPTS, MaxNumPolygons), ysv(MAXNUMPTS, MaxNumPolygons)
   ReDim Preserve xc(MaxNumPolygons), yc(MaxNumPolygons)
   ReDim Preserve SVNumPoints(MaxNumPolygons)
   ReDim Preserve LinHue(MaxNumPolygons)
   ReDim Preserve SaveSC(MaxNumPolygons), SaveEC(MaxNumPolygons)
   For k = 1 To NumPoints
      xsv(k, MaxNumPolygons) = xp(k)
      ysv(k, MaxNumPolygons) = yp(k)
   Next k
   xc(MaxNumPolygons) = X
   yc(MaxNumPolygons) = Y
   SVNumPoints(MaxNumPolygons) = NumPoints
   LinHue(MaxNumPolygons) = GradType
   SaveSC(MaxNumPolygons) = SColor
   SaveEC(MaxNumPolygons) = EColor
   cmdRUR(0).Enabled = True  ' Undo
   cmdRUR(2).Enabled = True  ' Delete
   NumPoints = 0   ' Signifies poly done
   LabTNum = Str$(PolygonNumber)
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show X,Y & Cross-hairs
   LabPNum = NumPoints
   LabXY = Str$(X) & "," & Str$(Y)
   If chkHairs Then
      With LineH
         .x1 = 0
         .y1 = Y
         .x2 = PIC.Width
         .y2 = Y
      End With
      With LineV
         .x1 = X
         .y1 = 0
         .x2 = X
         .y2 = PIC.Height
      End With
   End If
End Sub

Private Sub optGrad_Click(Index As Integer)
' 0 Linear, 1 Hue
   GradType = Index
End Sub

Private Sub optCenPt_Click(Index As Integer)
   If NumPoints <> 0 Then Exit Sub
   AutoManual = Index
End Sub

Private Sub cboRegular_Click()
Dim k As Long
   If NumPoints <> 0 Then Exit Sub
   k = cboRegular.ListIndex
   RegNumber = Val(cboRegular.List(k))
End Sub

Private Sub cboStarArm_Click()
Dim k As Long
   If NumPoints <> 0 Then Exit Sub
   k = cboStarArm.ListIndex
   zArmFrac = Val(cboStarArm.List(k))
End Sub

Private Sub chkRegular_Click()
   If NumPoints <> 0 Then Exit Sub
   aManualCenter = False
   aRegular = -chkRegular.Value
End Sub

Private Sub chkPoint_Click()
' Click on points ->
   If NumPoints <> 0 Then Exit Sub
   If aPlot Then
      chkPoint.Value = Unchecked
      aPlot = False
   Else
      chkPoint.Value = Checked
      aPlot = True
   End If
End Sub

Private Sub cmdSwap_Click(Index As Integer)
   If Index = 0 Then
      ' Swap SColor & EColor
      LabCul(0).BackColor = EColor
      LabCul(1).BackColor = SColor
      SColor = LabCul(0).BackColor
      EColor = LabCul(1).BackColor
   Else  ' Copy Start to End Color
      LabCul(1).BackColor = SColor
      EColor = LabCul(1).BackColor
   End If
   PIC.SetFocus
End Sub


Private Sub picCul_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LCul As Long
Dim bR As Byte, bG As Byte, bB As Byte
   LCul = picCul(Index).Point(X, Y)
   LabC(Index).BackColor = LCul
   LngToRGB LCul, bR, bG, bB
   With picRGB(Index)
      .Cls
      .ForeColor = vbRed
      picRGB(Index).Print Trim$(Str$(bR));
      .ForeColor = RGB(0, 120, 0)
      picRGB(Index).Print " " & Str$(bG);
      .ForeColor = vbBlue
      picRGB(Index).Print " " & Str$(bB);
   End With
End Sub

Private Sub picCul_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim svPolygonNumber As Long
Dim k As Long

   Select Case Index
   Case 0   ' Start color
      SColor = picCul(0).Point(X, Y)
      If SColor < 0 Then SColor = 0
      LabCul(0).BackColor = SColor
   Case 1   ' End color
      EColor = picCul(1).Point(X, Y)
      If EColor < 0 Then EColor = 0
      LabCul(1).BackColor = EColor
   Case 2   ' Back color
      If NumPoints = 0 Then
         BColor = picCul(1).Point(X, Y)
         If BColor < 0 Then BColor = 0
         PIC.BackColor = BColor
         LabCul(2).BackColor = BColor
         
         If MaxNumPolygons > 0 Then
            svPolygonNumber = PolygonNumber
            PolygonNumber = MaxNumPolygons
            DrawPolygons PolygonNumber
            PolygonNumber = svPolygonNumber
            NumPoints = SVNumPoints(PolygonNumber)
            For k = 1 To NumPoints
               xp(k) = xsv(k, PolygonNumber)
               yp(k) = ysv(k, PolygonNumber)
               shpPt(k - 1).Left = xp(k) - 2
               shpPt(k - 1).Top = yp(k) - 2
               shpPt(k - 1).Visible = True
            Next k
         End If
         NumPoints = 0
      End If
   End Select
End Sub

Private Sub chkHairs_Click()
' Cross-hairs
   If chkHairs Then
      LineH.Visible = True
      LineV.Visible = True
   Else
      LineH.Visible = False
      LineV.Visible = False
   End If
End Sub



Private Sub Form_Resize()
Dim PICR As Long, PICT As Long
   If WindowState = vbMinimized Then Exit Sub
   If Form1.Width < ORGFormWidth Or Form1.Height < ORGFormHeight Then
      Form1.Width = ORGFormWidth
      Form1.Height = ORGFormHeight
      Exit Sub
   End If
   PIC.Cls
   PICR = (Form1.Width) / STX - 50
   PICT = (Form1.Height) / STY - 100
   PIC.Width = PICR - PIC.Left
   PIC.Height = PICT - PIC.Top
   
   PIC.BackColor = BColor
   
   If MaxNumPolygons > 0 Then
      DrawPolygons MaxNumPolygons
      PolygonNumber = MaxNumPolygons
      cmdRUR(0).Enabled = True   ' Undo
      cmdRUR(1).Enabled = False  ' Redo
      cmdRUR(2).Enabled = True   ' Delete
      NumPoints = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   End
End Sub
