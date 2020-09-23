VERSION 5.00
Begin VB.Form FrmGrid 
   BackColor       =   &H80000001&
   Caption         =   "Editor Grid For Pictures"
   ClientHeight    =   8055
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14640
   Icon            =   "FrmGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Rotation options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   2700
      TabIndex        =   44
      Top             =   750
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   690
         TabIndex        =   54
         Text            =   "15"
         ToolTipText     =   "Enter an amount to incriment the paste by."
         Top             =   510
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Spin?"
         Height          =   255
         Left            =   1770
         TabIndex        =   53
         ToolTipText     =   "Paste the image many times in a circle."
         Top             =   210
         Width           =   765
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Rotate Forward?"
         Height          =   225
         Left            =   120
         TabIndex        =   52
         ToolTipText     =   "Rotate the image to be pasted forward if selected."
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Cancel"
         Height          =   285
         Index           =   1
         Left            =   450
         TabIndex        =   50
         Top             =   870
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "OK"
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   49
         Top             =   870
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2250
         TabIndex        =   46
         Text            =   "90"
         ToolTipText     =   "enter the amount of rotation or spin."
         Top             =   510
         Width           =   405
      End
      Begin VB.PictureBox selPoint 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   87
         TabIndex        =   45
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Current POINT:"
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
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   1830
         Width           =   2625
      End
      Begin VB.Label Label5 
         Caption         =   "Select a point in the image bellow to rotate around. or select ""OK"" to use the default, ""Center""."
         Height          =   615
         Left            =   150
         TabIndex        =   48
         Top             =   1200
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Enter the Degree of rotation."
         Height          =   225
         Left            =   60
         TabIndex        =   47
         Top             =   570
         Width           =   2925
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   6315
      Left            =   8070
      ScaleHeight     =   6255
      ScaleWidth      =   6435
      TabIndex        =   4
      Top             =   30
      Width           =   6495
      Begin VB.CommandButton Command4 
         Height          =   225
         Index           =   1
         Left            =   6210
         TabIndex        =   9
         Top             =   6030
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         LargeChange     =   100
         Left            =   0
         SmallChange     =   10
         TabIndex        =   7
         Top             =   6030
         Width           =   6195
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6015
         LargeChange     =   100
         Left            =   6210
         SmallChange     =   10
         TabIndex        =   6
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6015
         Left            =   0
         ScaleHeight     =   401
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   413
         TabIndex        =   5
         Top             =   0
         Width           =   6195
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   11490
            Left            =   0
            ScaleHeight     =   766
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   896
            TabIndex        =   8
            Top             =   0
            Width           =   13440
            Begin VB.Shape Shape1 
               BorderColor     =   &H80000005&
               BorderStyle     =   3  'Dot
               Height          =   780
               Left            =   0
               Shape           =   1  'Square
               Top             =   0
               Width           =   780
            End
         End
      End
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7995
      Left            =   30
      ScaleHeight     =   533
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   0
      Top             =   30
      Width           =   7995
      Begin VB.PictureBox PalPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7755
         Left            =   240
         ScaleHeight     =   517
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   514
         TabIndex        =   1
         Top             =   240
         Width           =   7710
      End
      Begin ImageEdit.rule rule1 
         Height          =   225
         Left            =   240
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   397
         SmallInterval   =   10
         CountType       =   1
         inVertit        =   -1  'True
         NumberInterval  =   100
         LargeInterval   =   50
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483634
      End
      Begin ImageEdit.rule rule2 
         Height          =   7755
         Left            =   0
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   13679
         SmallInterval   =   10
         CountType       =   1
         inVertit        =   -1  'True
         NumberInterval  =   100
         LargeInterval   =   50
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483634
         Orentation      =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   8070
      TabIndex        =   2
      Top             =   6390
      Width           =   6495
      Begin VB.CommandButton Command10 
         Caption         =   "&Copy / Move Tool"
         Height          =   285
         Index           =   3
         Left            =   3630
         TabIndex        =   40
         ToolTipText     =   "Display the copy move options."
         Top             =   30
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Grid &Options"
         Height          =   285
         Index           =   1
         Left            =   2430
         TabIndex        =   26
         ToolTipText     =   "Display the grid options (size,color)"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Grid Files"
         Height          =   285
         Index           =   0
         Left            =   1230
         TabIndex        =   23
         ToolTipText     =   "allows user to save and load map files."
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Edit &Tools"
         Height          =   285
         Index           =   2
         Left            =   30
         TabIndex        =   38
         ToolTipText     =   "displays the edit tools where you can choose a color or undo edits."
         Top             =   30
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Text            =   "64"
         Top             =   6660
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Height          =   1245
         Index           =   2
         Left            =   30
         TabIndex        =   32
         Top             =   300
         Width           =   6405
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1005
            Left            =   570
            Picture         =   "FrmGrid.frx":030A
            ScaleHeight     =   67
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   290
            TabIndex        =   37
            Top             =   150
            Width           =   4350
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Color Chooser"
            Height          =   255
            Left            =   4980
            TabIndex        =   36
            ToolTipText     =   "Choose a color to draw with"
            Top             =   180
            Width           =   1275
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   1005
            Left            =   30
            ScaleHeight     =   945
            ScaleWidth      =   465
            TabIndex        =   34
            Top             =   150
            Width           =   525
            Begin VB.PictureBox editPic 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   855
               Left            =   50
               ScaleHeight     =   57
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   35
               Top             =   50
               Width           =   375
            End
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Undo Last Edit"
            Height          =   285
            Left            =   4980
            TabIndex        =   33
            ToolTipText     =   "Undo the last edit"
            Top             =   750
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1245
         Index           =   3
         Left            =   30
         TabIndex        =   10
         Top             =   300
         Visible         =   0   'False
         Width           =   6405
         Begin VB.CommandButton Command2 
            Caption         =   "D,R"
            Height          =   255
            Index           =   6
            Left            =   1050
            TabIndex        =   20
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "D"
            Height          =   255
            Index           =   1
            Left            =   570
            TabIndex        =   14
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "R"
            Height          =   255
            Index           =   3
            Left            =   1050
            TabIndex        =   12
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "D,L"
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   21
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "U,R"
            Height          =   255
            Index           =   5
            Left            =   1050
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   5310
            TabIndex        =   16
            Text            =   "1"
            ToolTipText     =   "enter an amount by which you wnt the pixels to move."
            Top             =   870
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "U"
            Height          =   255
            Index           =   0
            Left            =   570
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "L"
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   13
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "U,L"
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "You must enable this tool to use it."
            Height          =   255
            Left            =   2130
            TabIndex        =   11
            Top             =   180
            Width           =   3915
         End
         Begin VB.Label Label3 
            Caption         =   "use buttons to move"
            Height          =   195
            Left            =   90
            TabIndex        =   39
            Top             =   990
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   $"FrmGrid.frx":E784
            Height          =   645
            Left            =   2100
            TabIndex        =   17
            Top             =   480
            Width           =   4185
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1245
         Index           =   1
         Left            =   30
         TabIndex        =   27
         Top             =   300
         Visible         =   0   'False
         Width           =   6405
         Begin VB.CommandButton Command7 
            Caption         =   "Change selection color"
            Height          =   285
            Index           =   1
            Left            =   1770
            TabIndex        =   43
            Top             =   600
            Width           =   1965
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Change grid color"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Draw New Grid"
            Height          =   285
            Left            =   3330
            TabIndex        =   30
            Top             =   120
            Width           =   1395
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2700
            TabIndex        =   29
            Text            =   "52"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Set grid size between 20 and 263"
            Height          =   255
            Left            =   90
            TabIndex        =   28
            Top             =   210
            Width           =   2505
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1245
         Index           =   0
         Left            =   30
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   6405
         Begin VB.CommandButton Command8 
            Caption         =   "Open a grid file you have saved with the "".map"" extension."
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   570
            Width           =   6015
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save any currently selected pixels to a map file you can use in any picture"
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   210
            Width           =   5955
         End
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu optOpen 
         Caption         =   "&Open"
         Begin VB.Menu Open 
            Caption         =   "Open &Picture"
         End
         Begin VB.Menu OpenGridFile 
            Caption         =   "Open &Map File"
         End
         Begin VB.Menu OpenBrush 
            Caption         =   "Open Patern &Brush"
         End
      End
      Begin VB.Menu optSave 
         Caption         =   "&Save"
         Begin VB.Menu save 
            Caption         =   "Save &Picture"
         End
         Begin VB.Menu SaveGridImage 
            Caption         =   "Save Grid &Image As Bmp"
         End
         Begin VB.Menu saveGridMap 
            Caption         =   "Save Grid As .&map"
         End
         Begin VB.Menu SaveBrush 
            Caption         =   "Save Pattern &Brush"
         End
      End
      Begin VB.Menu optView 
         Caption         =   "&View"
         Begin VB.Menu Grid 
            Caption         =   "View &Grid"
            Checked         =   -1  'True
         End
         Begin VB.Menu viewGridRule 
            Caption         =   "View Grid &Rule"
            Checked         =   -1  'True
         End
         Begin VB.Menu ViewCord 
            Caption         =   "View &Tool Tip Cord"
         End
         Begin VB.Menu ViewCross 
            Caption         =   "View &CrossHaire"
         End
      End
   End
   Begin VB.Menu Popup 
      Caption         =   "&Edit"
      Begin VB.Menu paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu pastFlipHor 
         Caption         =   "Paste Flip HOR"
      End
      Begin VB.Menu pasteFlipVer 
         Caption         =   "Paste Flip Ver"
      End
      Begin VB.Menu pasteRotate 
         Caption         =   "Paste Rotate"
      End
      Begin VB.Menu SelInside 
         Caption         =   "&Select Inside Outline"
      End
      Begin VB.Menu CreateBrush 
         Caption         =   "Create Pattern &Brush"
      End
      Begin VB.Menu SelectAllOf 
         Caption         =   "Select All Of: color"
      End
      Begin VB.Menu SelectAllBut 
         Caption         =   "Select All But: color"
      End
   End
   Begin VB.Menu UndoClear 
      Caption         =   "clear Undo Buffer"
   End
End
Attribute VB_Name = "FrmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mW As Long 'this is the size of the grid sections plus one for the border
Private selColor As Long 'current color selected
Private UnDoStk As Stack 'a dirty undo function
Private CD As cDlg
Private Const ScrCopy = &HCC0020 'used with stretchBlt
Private gridCol As Long 'allows the user to change the color of the grid lines
Private seleCol As Long ' this will allow the user to select a new highlight color 'not working
Private SelEnabl As Boolean 'allow the user to select certain pix and move them with the arrow keys
Private SelAry() As Long 'pixels that have been selected This is there points relative the main picture
Private GrdAry() As Long 'this is the selected pixels relative the selection square on the main picture
Private brush() As Long 'this will hold the pixels of a brush the brush should be
Private CurX As Long 'right click points on the picture grid
Private CurY As Long
Private oldX As Long 'used for drawing a selection
Private oldY As Long
Private pstX As Long 'used for the brush in mouse move to maintain the position of the start of that brush patern
Private pstY As Long
Private oldX1 As Long 'used for drawing a cross hair
Private oldY1 As Long
Private RuleY As Long
Private RuleX As Long
Private xc As Long 'used in the paste rotate sub
Private yc As Long
Private cancelRotat As Boolean
Private orgSel As Boolean 'used to determin if the user has moved the selection square since making first selection.
'*****  a new plan to allow the user to create a pattern brush from a selection in the picture grid **************
Private paternBrush As Boolean 'are we using a pattern brush
'*****************************************************************************************************************
Private makSel As Boolean 'tells a function if you are selecting with the right mouse down
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim R As RECT
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'testing for the shift key
Private Const VK_SHIFT = &H10      'during the selection process used for unselecting a pixel
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12

Private Sub Check1_Click() ' this is the selection option if it's checked we're selecting pixels
'if not we are coloring pixels.
    If Check1.Value = 1 Then
        SelEnabl = True 'set a veriable indicating we are selecting pixels
        ReDim SelAry(Shape1.Left To Val(Text2.Text) + Shape1.Left, Shape1.Top To Val(Text2.Text) + Shape1.Top) 'rebuild the arrays of selected points relative to the original picture
        ReDim GrdAry(Val(Text2.Text), Val(Text2.Text)) 'rebuild the array of selected points based on the selection grid.
        orgSel = True
    Else
        SelEnabl = False
        ShowSelection Shape1.Left, Shape1.Top
    End If
End Sub

Private Sub Check3_Click()
    If Check1.Value = 0 Then
    Label4.Caption = "Enter the Degree of rotation."
    Text4.Text = 90
    Text5.Visible = False
    Else
    Text5.Visible = True
    Label4.Caption = "Step By           deg.  Rotate to           deg"
    Text4.Text = 360
    End If
End Sub

Private Sub Command10_Click(Index As Integer) 'toolbox controls selecting one will hide the others and show it
Dim a As Long
    For a = 0 To 3
        Frame3(a).Visible = False
    Next a
Frame3(Index).Visible = True
End Sub

Private Sub Command8_Click() 'open a saved map file and select the pixels
'relative to the current location of the selection square
On Error Resume Next
Dim data As String
Dim parts() As String
Dim parts2() As String
Dim x As Long
Dim a As Long
Dim Y As Long
    CD.DefaultExt = "map"
    CD.FileName = ""
    CD.Filter = "Map Files(*.map)|*.map"
    CD.InitDir = App.Path & "\mapFile"
    CD.ShowOpen
        Open CD.FileName For Input As #1
            data = Input(LOF(1), 1)
        Close #1
    Text2.Text = Left(data, InStr(1, data, vbCrLf) - 1) 'set the grid size to what it was when this file was saved
    'Command5_Click 'redraw the grid to that size
    getGridSize Val(Text2.Text)
    Check1.Value = 0 'if there is something selected, unselect it
    DoEvents
    Check1.Value = 1 'start a new selection array
    data = Right(data, Len(data) - InStr(1, data, vbCrLf) + 1) 'remove the information from the string about grid size
    parts = Split(data, vbCrLf) 'split the remaining string into its x axes
        For x = 0 To UBound(parts) - 1 'loop through each x
            parts2 = Split(parts(x), ",") 'split the x into it's y axes
                For Y = 0 To UBound(parts2) - 1 'and loop through them
                    If CLng(parts2(Y)) <> 0 Then 'if the color is anything but 0(black)
'*******************************************************************************************************************
                'you could just write the pixel info to the picture now. I would rather be able to move my selection
                'before i paste it. If you want it pasted when you open it just uncomment the line bellow.
                        'SetPixel Picture3.hdc, Shape1.Left + x, Shape1.Top + Y, CLng(Parts2(Y))
'********************************************************************************************************************
                        SelAry(Shape1.Left + x, Shape1.Top + Y) = CLng(parts2(Y))
                        GrdAry(x, Y) = CLng(parts2(Y))
                        High_light_Pixel x, Y
                    End If
                Next Y
        Next x
    ShowSelection Shape1.Left, Shape1.Top 'refresh the grid to display the changes
End Sub

Private Sub Command1_Click() 'save a map file of the currently selected area
Dim data As String
Dim x As Long
Dim Y As Long
On Error Resume Next
    MkDir App.Path & "\Resources"
    MkDir App.Path & "\Resources\mapFiles"
        With CD
            '.FileName = ""
            .InitDir = App.Path & "\Resources\mapFiles\"
            .Filter = "Map Files(*.map)|*.map"
            .DefaultExt = "map"
            .ShowSave
        End With
    data = Text2.Text 'create the string that will be saved in the file. the first line is the grid size
        If SelEnabl = True Then 'if you are in selection mode
            For x = 0 To Shape1.Width 'loop through each line in the x cord.
            data = data & vbCrLf 'start a new line in the file for each new x
                For Y = 0 To Shape1.Height 'loop through each y in each x
                    If Y = 0 Then 'first item on a line
                        data = data & GrdAry(x, Y)
                    Else 'preseed each additional y with a comma that will be used to split the string
                        data = data & "," & GrdAry(x, Y)
                    End If
                Next Y
            Next x
        Else
            data = "" 'we dont have anything selected so exit.
            Exit Sub
            'optionally we could save all the pixels
        End If
        If CD.FileName <> "" Then
            Open CD.FileName For Output As #1 'save the file based on the name you selected
                Print #1, data                'at the start of the sub
            Close #1
        End If
End Sub

Private Sub Command2_Click(Index As Integer)
    MovePixel Index, Val(Text3.Text) 'call a sub to move the selected pixels "up,down,left, right or diag by an amount.)
End Sub

Private Sub MovePixel(Index As Integer, Optional distS As Long = 1) 'this is a small sub created to move any selected pixels in the direction
'indicated by an integer value up=0,dn=1,left=2,right=-3 the movement has to be done on the picture
'in the selected pixel array(SelAry), and the the Grid selected array(grdAry). if you move a selected pixel
'outside the area of the selection square those pixels will be unselected. also moves diagonaly
On Error Resume Next
    If orgSel = False Then Exit Sub
Dim Xx As Long          'used to indicate by how many pixels the selected pixels should move
     Xx = distS         'one is the default
Dim a As Long, B As Long
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
        Select Case Index
            Case 0 'Up
                For a = Shape1.Left To Shape1.Left + Shape1.Width 'look through the pixels incompassed by the
                For B = Shape1.Top To Shape1.Top + Shape1.Width   'selection square. For any color that isn't black(0)
                    If SelAry(a, B) <> 0 Then      'if you find one
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a, B - Xx, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a, B - Xx)
                        SetPixel Picture3.hdc, a, B - Xx, SelAry(a, B) 'draw that pixel onto the picture moving it the amount specified
                        'the bellow arrays' points are bsed on there position in the original picture and the bounds
                        SelAry(a, B - Xx) = SelAry(a, B) 'of the array is based on the position of the selection square on that picture
                        SelAry(a, B) = 0 'each time we move a pixel we must also move it's representitive pixel in SelAry.
                        'the GrdArys' points are based solely on the edit grid. It's size and colors are directly related to the selection square
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top)) - Xx) = SelAry(a, B - Xx) 'but the bounds of the array preset in that
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0 'the lBound will always be 0 and the ubound is always the width of the selection square
                        'the above pixel array will also need to be updated to the new location
                    End If
                Next B
                Next a
            Case 1 'Down
                For a = Shape1.Left To Shape1.Left + Shape1.Width
                For B = Shape1.Top + Shape1.Width To Shape1.Top Step -1
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a, B + Xx, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a, B + Xx)
                        SetPixel Picture3.hdc, a, B + Xx, SelAry(a, B)
                        SelAry(a, B + Xx) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top)) + Xx) = SelAry(a, B + Xx)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
            Case 2 'Left
                For a = Shape1.Left To Shape1.Left + Shape1.Width
                For B = Shape1.Top To Shape1.Top + Shape1.Width
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a - Xx, B, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a - Xx, B)
                        SetPixel Picture3.hdc, a - Xx, B, SelAry(a, B)
                        SelAry(a - Xx, B) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left) - Xx), ((B - Shape1.Top))) = SelAry(a - Xx, B)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
            Case 3 'Right
                For a = Shape1.Left + Shape1.Width To Shape1.Left Step -1
                For B = Shape1.Top To Shape1.Top + Shape1.Width
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a + Xx, B, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a + Xx, B)
                        SetPixel Picture3.hdc, a + Xx, B, SelAry(a, B)
                        SelAry(a + Xx, B) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left) + Xx), ((B - Shape1.Top))) = SelAry(a + Xx, B)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
            Case 4 'up and left
                For a = Shape1.Left To Shape1.Left + Shape1.Width
                For B = Shape1.Top To Shape1.Top + Shape1.Width
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a - Xx, B - Xx, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a - Xx, B - Xx)
                        SetPixel Picture3.hdc, a - Xx, B - Xx, SelAry(a, B)
                        SelAry(a - Xx, B - Xx) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left) - Xx), ((B - Shape1.Top)) - Xx) = SelAry(a - Xx, B - Xx)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
            Case 5 'up and right
                For a = Shape1.Left + Shape1.Width To Shape1.Left Step -1
                For B = Shape1.Top To Shape1.Top + Shape1.Width
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a + Xx, B - Xx, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a + Xx, B - Xx)
                        SetPixel Picture3.hdc, a + Xx, B - Xx, SelAry(a, B)
                        SelAry(a + Xx, B - Xx) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left) + Xx), ((B - Shape1.Top)) - Xx) = SelAry(a + Xx, B - Xx)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
            Case 6 'down and right
                For a = Shape1.Left + Shape1.Width To Shape1.Left Step -1
                For B = Shape1.Top + Shape1.Width To Shape1.Top Step -1
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a + Xx, B + Xx, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a + Xx, B + Xx)
                        SetPixel Picture3.hdc, a + Xx, B + Xx, SelAry(a, B)
                        SelAry(a + Xx, B + Xx) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left) + Xx), ((B - Shape1.Top)) + Xx) = SelAry(a + Xx, B + Xx)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
            Case 7 'down and left
                For a = Shape1.Left To Shape1.Left + Shape1.Width
                For B = Shape1.Top + Shape1.Width To Shape1.Top Step -1
                    If SelAry(a, B) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a - Xx, B + Xx, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a - Xx, B + Xx)
                        SetPixel Picture3.hdc, a - Xx, B + Xx, SelAry(a, B)
                        SelAry(a - Xx, B + Xx) = SelAry(a, B)
                        SelAry(a, B) = 0
                        GrdAry(((a - Shape1.Left) - Xx), ((B - Shape1.Top)) + Xx) = SelAry(a - Xx, B + Xx)
                        GrdAry(((a - Shape1.Left)), ((B - Shape1.Top))) = 0
                    End If
                Next B
                Next a
        End Select
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
    Picture3.Refresh
    ShowSelection Shape1.Left, Shape1.Top
End Sub

Public Sub HighLight_Selection() 'this sub is now called anytime check1.value = 1 each time check1s' value is 1
'the SelArys' and GrdArys' bounds are set. left mouse will redim those bounds based on the new position of the selection square
'right clicking however does not, it allows you to move the selection square to a new position for pasting
'therefore if you right click in the source when nothing is selected this sub would try to highlight pixels
'so we setup a couple of error handlers to save us from ourself.
On Error Resume Next '
Err.Clear 'insure there are no error messages left in the error buffer
    If UBound(GrdAry, 1) = 0 Then Exit Sub 'if there are no selected items
    If Err Then: Err.Clear: Exit Sub 'if we error getting the ubound of grdary
Dim a As Long, B As Long
    For a = 0 To Val(Text2.Text)
        For B = 0 To Val(Text2.Text)
            If GrdAry(a, B) <> 0 Then
                High_light_Pixel a, B
            End If
        Next B
    Next a
End Sub

Private Sub Command3_Click() 'choose a color to draw with uses the common dialog class
        With CD
            .Color = selColor
            .flags = cdlCCFullOpen Or cdlCCRGBInit Or cdlCCANYCOLOR 'allow custom colors for some reason it isn't saving or loading the colors properly.
            .ShowColor
        End With
    selColor = CD.Color
    editPic.BackColor = CD.Color
        If paternBrush = True Then
            paternBrush = False
            ReDim brush(0)
        End If
End Sub

Private Sub Command5_Click() 'rebulid the grid to the size indicated
    getGridSize Val(Text2.Text)
    ShowSelection Shape1.Left, Shape1.Top
    Picture3_MouseDown 1, 0, Shape1.Left, Shape1.Top 'reposition source if the selection square is outside the viewable area
End Sub

Private Sub Command6_Click() 'undo last edit this is only here so i could test my math
Dim Ret As VbMsgBoxResult
    If SelEnabl = True Then
        Ret = MsgBox("You are currently in selection mode and can't unselct with this method." & vbCrLf & "Would you like to continue with the undo?" & vbCrLf & "If you choose yes any selections you have made will be unselected, and the last edit reversed.", vbYesNo)
            If Ret = vbNo Then
                Exit Sub
            Else
                Check1.Value = 0
            End If
    End If
Dim Col As Long
Dim Y As Long
Dim x As Long
Dim newTop As Long
Dim NewLeft As Long
Dim srtx As Long
Dim srtY As Long
Dim a As Long
Dim B As Long
Dim UD() As Long
        If UnDoStk.stackLevel = 0 Then Exit Sub
    Screen.MousePointer = 11
    UD = UnDoStk.pop
    Col = UD(0)
    Y = UD(1)
    x = UD(2)
        If Not Picture3.Top = UD(3) Then 'we don't want to move the picture if we don't need to
            Picture3.Top = UD(3)
            Picture3.Left = UD(4)
        End If
        If Not Shape1.Top = UD(7) Or Not Shape1.Left = UD(8) Then 'we should move the shape if it's not in the right place
            getGridSize UD(5)
            Shape1.Move UD(8), UD(7), UD(6), UD(5)
            ShowSelection Shape1.Left, Shape1.Top
        End If
        If x = -1 Then 'key to indicate the start of a group of undos
            Do
                UD = UnDoStk.pop
                Col = UD(0) 'pop our info from the stack
                Y = UD(1)
                x = UD(2)
                newTop = UD(3)
                NewLeft = UD(4)
                    If x = -1 Then Exit Do 'key to indicate the end of the group
                srtx = x - Shape1.Left
                srtY = Y - Shape1.Top
                SetPixel Picture3.hdc, x, Y, Col
                Fill_Pixel srtx, srtY
            Loop
        Else
            srtx = x - Shape1.Left
            srtY = Y - Shape1.Top
            SetPixel Picture3.hdc, x, Y, Col
            Fill_Pixel srtx, srtY
        End If
    PalPic.Refresh
    Picture3.Refresh
    Screen.MousePointer = 0
End Sub

Private Sub Command7_Click(Index As Integer) 'allow the user to select a grid color. uses the common dialog
    With CD
        .flags = cdlCCFullOpen Or cdlCCRGBInit Or cdlCCANYCOLOR
        .ShowColor
    End With
    Select Case Index
        Case 0
            gridCol = CD.Color
            rule1.BackColor = CD.Color
            rule2.BackColor = CD.Color
            Command5_Click
        Case 1
            seleCol = CD.Color
            Shape1.BorderColor = seleCol
            rule1.ForeColor = CD.Color
            rule2.ForeColor = CD.Color
            ShowSelection Shape1.Left, Shape1.Top
    End Select
End Sub

Private Sub Command9_Click(Index As Integer)
Dim Ret As VbMsgBoxResult
    If Index = 1 Then
        cancelRotat = True
    Else
        If Val(Text4.Text) > 360 Then
            Ret = MsgBox("You have entered a value greater than 360 deg." & vbCrLf & "It has ben reset to 360." & vbCrLf & vbCrLf & "If you want to reset it press " & """" & "Cancel" & """", vbOKCancel)
                If Ret = vbCancel Then
                    Exit Sub
                End If
        End If
    End If
Frame2.Visible = False
End Sub

Private Sub CreateBrush_Click()
'this option will be used to draw with a patern brush. (pattern determined by the pixels selected)
'this only creates square brushes. and the size is determined by the number of selected pixels in a row
'in the y direction.
    If Check1.Value = 0 Then
        Check1.Value = 1
        Command10_Click 3
        MsgBox "select the area you want to use to create a brush" & vbCrLf & "Your selection must be square." & vbCrLf & vbCrLf & "Then unselect the " & """" & "Enable Select Tool" & """" & " option."
            Do
                Sleep 1
                DoEvents
            Loop Until Check1.Value = 0
    End If
Dim x As Long
Dim Y As Long
Dim srt As Boolean
Dim pntX As Long 'the x cord of the first selected pixel
Dim pntY As Long 'the y cord of first selected pixel
Dim wid As Long 'the number of selected pixels in the first colum
    wid = 0
    srt = False
        For x = 0 To Shape1.Width 'loop through each pixel
            For Y = 0 To Shape1.Height
                If GrdAry(x, Y) <> 0 Then 'if the corasponding pixel in the grid array is anything but 0
                        If srt = False Then 'if this is the first selected pixel found
                            pntX = x 'load up the starting x,y
                            pntY = Y
                        End If
                    srt = True 'we are still finding pixels on in this colum
                Else
                    srt = False 'we have found the last selected pixel
                End If
                If srt = True Then 'if we are finding pixels
                    wid = wid + 1 'incriment the count by one
                ElseIf Not wid = 0 Then 'if this pixel isn't selected and we have at least one found
                    'set the height and width of the brush
                    ReDim brush(pntX To pntX + wid, pntY To pntY + wid)
                    'once you have determined the number of pixels selected in the first colum just make the width the same size
                    GoTo getout
                End If
            Next Y
        Next x
'if we make it to here we haven't selected anything yet
    MsgBox "You must select a vertical row of pixels to represent the width and height of the Brush."
    Exit Sub
getout:
        For x = pntX To pntX + wid - 1
            For Y = pntY To pntY + wid - 1
                brush(x, Y) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top) '  GrdAry(X, Y)
            Next Y
        Next x
    paternBrush = True
    Check1.Value = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 'use the arrow keys to move pixels if the
'"enable select tool" is active and move the selection square if not.
Dim a As Long, B As Long
        If KeyCode = 16 Then Exit Sub
        If Shift = 2 Then
            Select Case KeyCode
                Case 90
                     Command6_Click
                Case 69 'E
                    Command10_Click 2
                Case 71 'G
                    Command10_Click 0
                Case 79 'O
                    Command10_Click 1
                Case 67 'C
                    Command10_Click 3
                        If Check1.Value = 1 Then
                            Check1.Value = 0
                        Else
                            Check1.Value = 1
                        End If
            End Select
        End If
        If SelEnabl = False Then 'if we arnt selecting items to move then move the selection square
            Select Case KeyCode
                Case 38 'Up
                        If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then 'move the selection square by its' height
                            If Shape1.Top - Shape1.Height >= 0 Then
                                Shape1.Top = Shape1.Top - Shape1.Height
                                If Picture3.Top < -Shape1.Top Then Picture3.Top = -Shape1.Top
                            Else
                                Shape1.Top = 0
                                Picture3.Top = 0
                            End If
                        Else 'move by 1
                            If Shape1.Top > 0 Then Shape1.Top = Shape1.Top - 1
                            If -Picture3.Top > Shape1.Top Then Picture3.Top = Picture3.Top + 1
                        End If
                    VScroll1.Value = -Picture3.Top
                    ShowSelection Shape1.Left, Shape1.Top
                Case 40 'Down
                        If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                            If Shape1.Top + (Shape1.Height * 2) < Picture3.ScaleHeight Then
                                Shape1.Top = Shape1.Top + Shape1.Height
                                If Shape1.Top + Shape1.Height > (-Picture3.Top + Picture5.ScaleHeight) Then Picture3.Top = -((Shape1.Top + Shape1.Height) - Picture5.ScaleHeight)
                            Else
                                Shape1.Top = Picture3.ScaleHeight - Shape1.Height
                                Picture3.Top = -(Picture3.ScaleHeight - Picture5.ScaleHeight)
                            End If
                        Else
                            If Shape1.Top < Picture3.ScaleHeight - Shape1.Height Then Shape1.Top = Shape1.Top + 1
                            If -Picture3.Top + Picture5.ScaleHeight < Shape1.Top + Shape1.Height Then Picture3.Top = Picture3.Top - 1
                        End If
                    VScroll1.Value = -Picture3.Top
                    ShowSelection Shape1.Left, Shape1.Top
                Case 37 'Left
                        If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                            If Shape1.Left - Shape1.Width >= 0 Then
                                Shape1.Left = Shape1.Left - Shape1.Width
                                    If Picture3.Left < -Shape1.Left Then Picture3.Left = -Shape1.Left
                            Else
                                Shape1.Left = 0
                                Picture3.Left = 0
                            End If
                        Else
                            If Shape1.Left > 0 Then Shape1.Left = Shape1.Left - 1
                            If -Picture3.Left > Shape1.Left Then Picture3.Left = Picture3.Left + 1
                        End If
                    HScroll1.Value = -Picture3.Left
                    ShowSelection Shape1.Left, Shape1.Top
                Case 39 'Right
                        If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                            If Shape1.Left + (Shape1.Width * 2) < Picture3.ScaleWidth Then
                                Shape1.Left = Shape1.Left + Shape1.Width
                                    If Shape1.Left + Shape1.Width > (-Picture3.Left + Picture5.ScaleHeight) Then Picture3.Left = -((Shape1.Left + Shape1.Width) - Picture5.ScaleWidth)
                            Else
                                Shape1.Left = Picture3.ScaleWidth - Shape1.Width
                                Picture3.Left = -(Picture3.ScaleWidth - Picture5.ScaleWidth)
                            End If
                        Else
                            If Shape1.Left < Picture3.ScaleWidth - Shape1.Width Then Shape1.Left = Shape1.Left + 1
                            If -Picture3.Left + Picture5.ScaleWidth < Shape1.Left + Shape1.Width Then Picture3.Left = Picture3.Left - 1
                        End If
                    HScroll1.Value = -Picture3.Left
                    ShowSelection Shape1.Left, Shape1.Top
            End Select
        Else
            Select Case KeyCode
                Case 38 'Up
                    MovePixel 0, Val(Text3.Text)
                Case 40 'Down
                    MovePixel 1, Val(Text3.Text)
                Case 37 'Left
                    MovePixel 2, Val(Text3.Text)
                Case 39 'Right
                    MovePixel 3, Val(Text3.Text)
            End Select
        End If
End Sub

Private Sub Form_Load()
    makSel = False 'initial settings  ************
    getGridSize CLng(GetSetting(App.EXEName, "Settings", "GridSize", "52")) 'load the grid veriables to programs' memory
    seleCol = GetSetting(App.EXEName, "Settings", "Highlight", 16777215) 'highlight color
    selColor = GetSetting(App.EXEName, "Settings", "DrawingColor", 0)
    gridCol = GetSetting(App.EXEName, "Settings", "gridcolor", 0)
    rule1.ForeColor = seleCol: rule2.ForeColor = seleCol: Shape1.BorderColor = seleCol
    rule1.BackColor = gridCol: rule2.BackColor = gridCol
    editPic.BackColor = GetSetting(App.EXEName, "Settings", "Backcolor", 0)
    Set UnDoStk = New Stack: Set UnDoStk.desp = UndoClear: UnDoStk.ClearUndo  'setup the undo class
    Set CD = New cDlg: CD.hOwner = Me.hWnd 'setup common dialog class
'    Picture3_Resize 'force a refresh of the controls
'    ShowSelection 0, 0
'reset Last Position and size
Dim l As Long, t As Long, W As Long, H As Long
    l = GetSetting(App.EXEName, "Settings", "MainLeft", 1000)
    t = GetSetting(App.EXEName, "Settings", "MainTop", 1000)    'set the window to the last user size
    W = GetSetting(App.EXEName, "Settings", "MainWidth", 15645)    'if the window is restored from a max window sate
    H = GetSetting(App.EXEName, "Settings", "MainHeight", 8145)    'it will be returned to this size
    Me.Move l, t, W, H
        If GetSetting(App.EXEName, "Settings", "MaxWindow", "True") = True Then
            Me.WindowState = 2 'if the window was maxamized then reset its position
        End If
'reset options
    Me.Grid.Checked = GetSetting(App.EXEName, "Settings", "Grid", True)
    Me.viewGridRule.Checked = GetSetting(App.EXEName, "Settings", "Rule", True)
    Me.ViewCord.Checked = GetSetting(App.EXEName, "Settings", "ToolTips", False)
    Me.ViewCross.Checked = GetSetting(App.EXEName, "Settings", "CrosHar", False)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If Not Me.WindowState = 2 Then
            Call SaveSetting(App.EXEName, "Settings", "MainLeft", Me.Left)
            Call SaveSetting(App.EXEName, "Settings", "MainTop", Me.Top)
            Call SaveSetting(App.EXEName, "Settings", "MainWidth", Me.Width)
            Call SaveSetting(App.EXEName, "Settings", "MainHeight", Me.Height)
        End If
    Call SaveSetting(App.EXEName, "Settings", "MaxWindow", Me.WindowState = 2)
    Call SaveSetting(App.EXEName, "Settings", "Grid", Me.Grid.Checked)
    Call SaveSetting(App.EXEName, "Settings", "Rule", Me.viewGridRule.Checked)
    Call SaveSetting(App.EXEName, "Settings", "ToolTips", Me.ViewCord.Checked)
    Call SaveSetting(App.EXEName, "Settings", "CrosHar", Me.ViewCross.Checked)
    Call SaveSetting(App.EXEName, "Settings", "Highlight", CStr(seleCol))
    Call SaveSetting(App.EXEName, "Settings", "DrawingColor", CStr(selColor))
    Call SaveSetting(App.EXEName, "Settings", "GridSize", CStr(Val(Text2.Text) - 5))
    Call SaveSetting(App.EXEName, "Settings", "gridcolor", CStr(gridCol))
    Call SaveSetting(App.EXEName, "Settings", "Backcolor", CStr(editPic.BackColor))

End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
            If Me.Width < (Me.Height + Frame1.Width) Then
                    If Not Me.WindowState = 2 Then
                        Me.Width = (Me.Height + Frame1.Width)
                    Else
                        PicMain.Move 30, 30, Me.ScaleWidth - 60 - Frame1.Width, Me.ScaleWidth - 60 - Frame1.Width
                        Picture4.Move PicMain.Left + PicMain.Width + 40, 30, Me.ScaleWidth - (PicMain.Left + PicMain.Width + 60), Me.ScaleHeight - Frame1.Height - 90
                        Frame1.Move Picture4.Left, Picture4.Top + Picture4.Height + 30, Frame1.Width, Frame1.Height
                    End If
                Exit Sub
            End If
        PicMain.Move 30, 30, Me.ScaleHeight - 60, Me.ScaleHeight - 60
        Picture4.Move PicMain.Left + PicMain.Width + 40, 30, Me.ScaleWidth - (PicMain.Left + PicMain.Width + 60), Me.ScaleHeight - Frame1.Height - 90
        Frame1.Move Picture4.Left, Picture4.Top + Picture4.Height + 30, Frame1.Width, Frame1.Height
    End If
End Sub

Private Sub Grid_Click() 'switch between grid or no grid
    Grid.Checked = Not Grid.Checked 'toggle grid
    ShowSelection Shape1.Left, Shape1.Top 'refresh grid with new grid settings
End Sub

Private Sub HScroll1_Change()
    Picture3.Left = -HScroll1.Value 'match the pictures' position with the value of the scrollbar
End Sub

Private Sub HScroll1_GotFocus()
    Picture3.SetFocus 'stop th scrollbar from blinking
End Sub

Private Sub HScroll1_Scroll()
    Picture3.Left = -HScroll1.Value
End Sub

Private Sub Open_Click()
On Error Resume Next
    With CD
        .Filter = "Supported Pictures(*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|BitMap(*.bmp)|*.bmp|Gif(*.gif)|*.gif|Jpeg (*.jpg)|*.jpg|Icon(*.ico)|*.ico|All Files(*.*)|*.*"
        .ShowOpen
    End With
    If LCase(Right(CD.FileName, 3)) = "lnk" Then
        MsgBox "This is not a valid folder or file"
    Else
            If CD.FileName <> "" Then Picture3.Picture = LoadPicture(CD.FileName)
        getGridSize Val(Text2.Text)
        ShowSelection Shape1.Left, Shape1.Top 'refresh grid
        UnDoStk.ClearUndo
    End If
End Sub

Private Sub OpenGridFile_Click()
    Command8_Click
End Sub

Private Sub PalPic_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    If pasting = True Then
'        pasting = False
'        xc = x \ mW: yc = Y \ mW
'        Exit Sub
'    End If
Dim srtx As Long
Dim srtY As Long
Dim a As Long, B As Long
    srtx = x \ (mW)
    srtY = Y \ (mW)
    oldX = x
    oldY = Y
        If SelEnabl = True Then
            If Button = 2 Then
                CurX = (x \ mW) + Shape1.Left
                CurY = (Y \ mW) + Shape1.Top
            ElseIf Button = 1 Then
    'Need to check the bounds of the array to see if this pixel is ouside that selection
    'if the selection is outside pop a message telling the user to enlarge the selection square.
    'or make two selections.
                If orgSel = True Then
                    If x Mod mW > 0 And Y Mod mW > 0 Then
                        If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                            High_light_Pixel srtx, srtY, True
                            SelAry(srtx + Shape1.Left, srtY + Shape1.Top) = 0
                            GrdAry(srtx, srtY) = 0
                        Else
                            High_light_Pixel srtx, srtY
                            SelAry(srtx + Shape1.Left, srtY + Shape1.Top) = GetPixel(PalPic.hdc, x, Y)
                            GrdAry(srtx, srtY) = GetPixel(PalPic.hdc, x, Y)
                        End If
                    End If
                Else
                    'give the user an option to view this error in the feuture.
                    MsgBox "After moving the selection square, You can't choose more pixels." & vbCrLf & "If you need to select these pixels, You will need to enlarge your selection square to incompass them." & vbCrLf & "You can do this by right clicking and draging to the desired size, or enter a larger value" & vbCrLf & "in the grid options."
                End If
            End If
        Else
            If Button = 2 Then
'                'right click select color was moved to mouse up
            ElseIf Button = 1 Then
                If paternBrush = True Then
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'Push an undo loop indicator
                        For a = 0 To (UBound(brush, 1) - LBound(brush, 1)) - 1
                            For B = 0 To (UBound(brush, 2) - LBound(brush, 2)) - 1
                                UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B)
                                SetPixel Picture3.hdc, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B, brush(a + LBound(brush, 1), B + LBound(brush, 2))
                            Next B
                        Next a
                    pstX = (x \ mW) + Shape1.Left
                    pstY = (Y \ mW) + Shape1.Top
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'end the loop indicator
                    Picture3.Refresh
                    ShowSelection Shape1.Left, Shape1.Top
                Else
                        If GetPixel(PalPic.hdc, ((srtx * mW) + 1), ((srtY * mW) + 1)) = selColor Then Exit Sub 'no need to color the point the same color
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, srtx + Shape1.Left, srtY + Shape1.Top, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, srtx + Shape1.Left, srtY + Shape1.Top)
                    SetPixel Picture3.hdc, srtx + Shape1.Left, srtY + Shape1.Top, selColor
                    Picture3.Refresh
                    Fill_Pixel srtx, srtY
                    PalPic.Refresh
                End If
            End If
        End If
End Sub

Private Sub PalPic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If ViewCord.Checked = True Then
        PalPic.ToolTipText = (x \ mW) + Shape1.Left & "," & (Y \ mW) + Shape1.Top
    Else
        PalPic.ToolTipText = ""
    End If
Dim srtx As Long, a As Long, B As Long
Dim srtY As Long
    If viewGridRule.Checked = True Then
        rule1.PointOn Int(x), Int(0), vbYellow
        rule2.PointOn 0, Int(Y), vbYellow
    End If
On Error Resume Next
    If SelEnabl = True Then
        If Button = 1 Then 'holding the left mouse then highlight each pixel as we move over it
            If x Mod mW > 0 And Y Mod mW > 0 Then
                srtx = x \ (mW)
                srtY = Y \ (mW)
                    If orgSel = True Then 'precaution aginst the user selecting pixels in one location, moving the selection square, and trying to select pixels in the new area.
                        If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then 'shift key is down
                            High_light_Pixel srtx, srtY, True
                            SelAry(srtx + Shape1.Left, srtY + Shape1.Top) = 0
                            GrdAry(srtx, srtY) = 0
                        Else
                            High_light_Pixel srtx, srtY
                            SelAry(srtx + Shape1.Left, srtY + Shape1.Top) = GetPixel(PalPic.hdc, x, Y)
                            GrdAry(srtx, srtY) = GetPixel(PalPic.hdc, x, Y)
                        End If
                    End If
            End If
        ElseIf Button = 2 Then 'right mouse to select an area to highlight
            If x <> oldX Or Y <> oldY Then
                If makSel = True Then DrawFocusRect PalPic.hdc, R 'color out any previous drawing
                    makSel = True
                        If Y > oldY And x > oldX Then
                            R.Bottom = Y
                            R.Left = oldX
                            R.Right = x
                            R.Top = oldY
                        ElseIf Y < oldY And x > oldX Then
                            R.Bottom = oldY
                            R.Left = oldX
                            R.Right = x
                            R.Top = Y
                        ElseIf Y > oldY And x < oldX Then
                            R.Bottom = Y
                            R.Left = x
                            R.Right = oldX
                            R.Top = oldY
                        ElseIf Y < oldY And x < oldX Then
                            R.Bottom = oldY
                            R.Left = x
                            R.Right = oldX
                            R.Top = Y
                        End If
                    DrawFocusRect PalPic.hdc, R 'draw to the new position.
                    PalPic.Refresh
            End If
        End If
    Else
        If Button = 1 Then
            If paternBrush = True Then
'***************************************************************************************************************************************************************************************************************
'Testing the pattern brush. The loops bellow test to see if the mouse has moved the width or height in the + or -
'direction. If the result is true another patern is drawn. This prevents streaking of the colors. It is basicly
'just tiling the pattern based on the mouse position.
                If pstX + (UBound(brush, 1) - LBound(brush, 1)) = ((x \ mW) + Shape1.Left) Or pstY + (UBound(brush, 2) - LBound(brush, 2)) = (Y \ mW) + Shape1.Top Then
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'Push an undo loop indicator
                        For a = 0 To (UBound(brush, 1) - LBound(brush, 1)) - 1
                            For B = 0 To (UBound(brush, 2) - LBound(brush, 2)) - 1
                                UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B)
                                SetPixel Picture3.hdc, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B, brush(a + LBound(brush, 1), B + LBound(brush, 2))
                            Next B
                        Next a
                    pstX = (x \ mW) + Shape1.Left '+ A
                    pstY = (Y \ mW) + Shape1.Top '+ B
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'end the loop indicator
                    Picture3.Refresh
                    ShowSelection Shape1.Left, Shape1.Top
                ElseIf pstX - (UBound(brush, 1) - LBound(brush, 1)) = ((x \ mW) + Shape1.Left) Or pstY - (UBound(brush, 2) - LBound(brush, 2)) = (Y \ mW) + Shape1.Top Then
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'Push an undo loop indicator
                        For a = 0 To (UBound(brush, 1) - LBound(brush, 1)) - 1
                            For B = 0 To (UBound(brush, 2) - LBound(brush, 2)) - 1
                                UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B)
                                SetPixel Picture3.hdc, (x \ mW) + Shape1.Left + a, (Y \ mW) + Shape1.Top + B, brush(a + LBound(brush, 1), B + LBound(brush, 2))
                            Next B
                        Next a
                    pstX = (x \ mW) + Shape1.Left '+ A
                    pstY = (Y \ mW) + Shape1.Top '+ B
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'end the loop indicator
                    Picture3.Refresh
                    ShowSelection Shape1.Left, Shape1.Top
                End If
'*****************************************************************************************************************************************************************************************************************
            Else
                srtx = x \ (mW)
                srtY = Y \ (mW)
                    If GetPixel(PalPic.hdc, ((srtx * mW) + 1), ((srtY * mW) + 1)) = selColor Then Exit Sub 'no need to color the point the same color
                UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, srtx + Shape1.Left, srtY + Shape1.Top, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, srtx + Shape1.Left, srtY + Shape1.Top)
                SetPixel Picture3.hdc, srtx + Shape1.Left, srtY + Shape1.Top, selColor
                Picture3.Refresh
                Fill_Pixel srtx, srtY
                PalPic.Refresh
            End If
        ElseIf Button = 2 Then
            'Draw A Focus Rec around the pixels
            'allowing you to color them at one time
            If x <> oldX Or Y <> oldY Then
                DrawFocusRect PalPic.hdc, R
                makSel = True
                    If Y > oldY And x > oldX Then
                        R.Bottom = Y
                        R.Left = oldX
                        R.Right = x
                        R.Top = oldY
                    ElseIf Y < oldY And x > oldX Then
                        R.Bottom = oldY
                        R.Left = oldX
                        R.Right = x
                        R.Top = Y
                    ElseIf Y > oldY And x < oldX Then
                        R.Bottom = Y
                        R.Left = x
                        R.Right = oldX
                        R.Top = oldY
                    ElseIf Y < oldY And x < oldX Then
                        R.Bottom = oldY
                        R.Left = x
                        R.Right = oldX
                        R.Top = Y
                    End If
                DrawFocusRect PalPic.hdc, R
                PalPic.Refresh
            End If
        Else
        End If
    End If
'*****************************************Draw Cross Hairs*******************************************************
'this still needs work.
'I don't like this at all; but, left it for you to see and test with. You might be able to make it work. The thought
'was to draw a line across the picture along the X and Y of the mouse. My thought was to align these new lines with
'the grid. The four lines commented out do that; however there are conflicts with the other drawing going on. If
'you are selecting pixels in the grid the selection lines drawn around the picture are affected. If there isn't a
'grid viewable the edges of the pixel are not colored right. Therefore I just moved the lines outside the pixels'
'bounderies. but there are problems with that as well. The only thing I have come up with to fix the error is to
'remove the cross hairs at the mouse down event and redraw at mouse up, but I haven't tested it yet. If you find a
'fix let me know.
            If ViewCross.Checked = True Then
                PalPic.DrawMode = 7 'XorPen
'                If oldX1 <> 0 Then PalPic.Line (oldX1 - (oldX1 Mod mW), 0)-(oldX1 - (oldX1 Mod mW), PalPic.ScaleHeight), vbWhite
'                PalPic.Line (X - (X Mod mW), 0)-(X - (X Mod mW), PalPic.ScaleHeight), vbWhite
'                If oldY1 <> 0 Then PalPic.Line (0, oldY1 - (oldY1 Mod mW))-(PalPic.ScaleWidth, oldY1 - (oldY1 Mod mW)), vbWhite
'                PalPic.Line (0, Y - (Y Mod mW))-(PalPic.ScaleWidth, Y - (Y Mod mW)), vbWhite
                
                If oldX1 <> 0 Then PalPic.Line ((oldX1 - (oldX1 Mod mW)) - 1, 0)-((oldX1 - (oldX1 Mod mW)) - 1, PalPic.ScaleHeight), vbWhite
                PalPic.Line ((x - (x Mod mW)) - 1, 0)-((x - (x Mod mW)) - 1, PalPic.ScaleHeight), vbWhite
                If oldY1 <> 0 Then PalPic.Line (0, (oldY1 - (oldY1 Mod mW)) - 1)-(PalPic.ScaleWidth, (oldY1 - (oldY1 Mod mW)) - 1), vbWhite
                PalPic.Line (0, (Y - (Y Mod mW)) - 1)-(PalPic.ScaleWidth, (Y - (Y Mod mW)) - 1), vbWhite
                
                PalPic.DrawMode = 13 'copyPen
                oldX1 = x: oldY1 = Y
            End If
'****************************************************************************************************************
End Sub

Private Sub PalPic_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim numPartsX As Long
Dim numPartsY As Long
Dim srtx As Long
Dim srtY As Long
Dim a As Long, B As Long
    If makSel = True Then
            DrawFocusRect PalPic.hdc, R 'clear any rect we have drawn
            PalPic.Refresh
            numPartsX = (R.Right - R.Left) \ mW '
            srtx = R.Left \ mW
            srtY = R.Top \ mW
            numPartsY = (R.Bottom - R.Top) \ mW
                If SelEnabl = False Then UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
                For a = srtx To numPartsX + srtx
                    For B = srtY To numPartsY + srtY
                        If SelEnabl = True Then
                            If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                                High_light_Pixel a, B, True
                                SelAry(a + Shape1.Left, B + Shape1.Top) = 0
                                GrdAry(a, B) = 0
                            Else
                                High_light_Pixel a, B
                                SelAry(a + Shape1.Left, B + Shape1.Top) = GetPixel(PalPic.hdc, (a * mW) + 1, (B * mW) + 1)
                                GrdAry(a, B) = GetPixel(PalPic.hdc, (a * mW) + 1, (B * mW) + 1)
                            End If
                        Else
                            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, a + Shape1.Left, B + Shape1.Top, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, a + Shape1.Left, B + Shape1.Top)
                            SetPixel Picture3.hdc, a + Shape1.Left, B + Shape1.Top, selColor
                            Fill_Pixel a, B
                        End If
                    Next B
                Next a
                If SelEnabl = False Then UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
            Picture3.Refresh
            PalPic.Refresh
            R.Bottom = -1: R.Left = -1: R.Right = -1: R.Top = -1
        makSel = False
    ElseIf SelEnabl = True Then
            If Button = 2 Then
                CurX = (x \ mW) + Shape1.Left
                CurY = (Y \ mW) + Shape1.Top
                PopupMenu Popup
            Else
            End If
    Else
        If Button = 2 Then
            If x Mod mW > 0 And Y Mod mW > 0 Then
                selColor = GetPixel(PalPic.hdc, x, Y)
                editPic.BackColor = GetPixel(PalPic.hdc, x, Y)
            End If
                If paternBrush = True Then
                    paternBrush = False
                    ReDim brush(0)
                End If
        End If
    End If
End Sub

Private Sub Paste_Click()
'simply loop through the array of points and color. the points are based from the edges of the
'selection square. Wherever your selections are when you choose it in the edit window.
Dim a As Long
Dim B As Long
On Error Resume Next
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
        For a = 0 To Shape1.Width
            For B = 0 To Shape1.Height
                If GrdAry(a, B) <> 0 Then
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + (a), Shape1.Top + (B), Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, Shape1.Left + (a), Shape1.Top + (B)) 'key to start a loop of undo events
                    SetPixel Picture3.hdc, Shape1.Left + (a), Shape1.Top + (B), GrdAry(a, B)
                End If
            Next B
        Next a
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
    Picture3.Refresh
    ShowSelection Shape1.Left, Shape1.Top
End Sub

Private Sub pasteFlipVer_Click() 'swap the top pixels to the bottom
Dim a As Long
Dim B As Long
On Error Resume Next
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
        For a = 0 To Shape1.Width
            For B = 0 To Shape1.Height
                If GrdAry(a, B) <> 0 Then
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + (a), (Shape1.Top + (Shape1.Height - B)), Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, Shape1.Left + (a), (Shape1.Top + (Shape1.Height - B)))
                    SetPixel Picture3.hdc, Shape1.Left + (a), (Shape1.Top + (Shape1.Height - B)), GrdAry(a, B)
                End If
            Next B
        Next a
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
    Picture3.Refresh
    ShowSelection Shape1.Left, Shape1.Top
End Sub

Private Sub pasteRotate_Click() 'the rotation is counter clockwise
'this is a rotation function that was created to rotate an image based on a point in that image. It was
'good code and worked well. I have totally scrued it up. trying to get it to rotate a portion of a picture
'with a center origen It is now working to rotate the image but pixels are being lost in the process at certain angles.
'
Const PI As Single = 3.141592653
Dim radians As Single
Dim angle As Single, angle0 As Single
Dim distance As Single
Dim rotFor As Boolean
Dim deltaX As Long, deltaY As Long
Dim x As Long, Y As Long
Dim x0 As Long, y0 As Long
'Dim xc As Long, yc As Long 'these are the point in the image being rotated that we are rotating around.
Dim degrees As Single
On Error Resume Next
'***************************Testing a rotation option dialog*********************************************************
'xc and yc must be declared form wide allowing us to set them in the options dialog
'initially they should be set to center.
    xc = Shape1.Width \ 2 'this rotates from the center of the grid. This is where the problem was in this
    yc = Shape1.Height \ 2 'sub I never set them to the center.
    selPoint.Move 120, 2130, Shape1.Width * Screen.TwipsPerPixelX, Shape1.Height * Screen.TwipsPerPixelY
    Frame2.Height = selPoint.Top + selPoint.Height + 120
    Frame2.Width = 3015
    Label6.Caption = "POINT:" & xc & "," & yc
    selPoint.Cls
        If Frame2.Width < selPoint.Width + 120 Then Frame2.Width = selPoint.Width + 120
    Frame2.Visible = True: cancelRotat = False
        For x = 0 To Shape1.Width 'draw the selected image into our rotation dialog
            For Y = 0 To Shape1.Height
                If GrdAry(x, Y) <> 0 Then
                    Call SetPixel(selPoint.hdc, x, Y, GrdAry(x, Y))
                End If
            Next Y
        Next x
    selPoint.Refresh
        Do 'loop here until the user selects ok or cancle
            DoEvents
            Sleep 1
        Loop Until Frame2.Visible = False
        If cancelRotat = True Then Exit Sub
        If Check3.Value = 0 Then 'if we are using the new spin option
            degrees = Val(Text4.Text)
        Else
            degrees = 0
        End If
    rotFor = Check2.Value
'*******************************************************************************************************************
        If degrees < 0 Then degrees = 0
        If degrees > 360 Then degrees = 360
redo:
    radians = degrees / (180 / PI)
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
        For x = 0 To Shape1.Width
            For Y = 0 To Shape1.Height
                deltaX = x - xc
                deltaY = Y - yc
                    If deltaX > 0 Then
                        angle = Atn(deltaY / deltaX)
                    ElseIf deltaX < 0 Then
                        angle = PI + Atn(deltaY / deltaX)
                    Else
                        If deltaY > 0 Then angle = PI / 2 Else angle = PI * 3 / 2
                    End If
                    If rotFor = True Then 'rotate forward
                        angle0 = angle + radians
                    Else 'rotate back
                        angle0 = angle - radians
                    End If
                distance = Sqr(deltaX * deltaX + deltaY * deltaY)
                x0 = xc + distance * Cos(angle0)
                y0 = yc + distance * Sin(angle0)
                    If GrdAry(x, Y) <> 0 Then
                        UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, (Shape1.Left + x0), (Shape1.Top + y0), Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, (Shape1.Left + x0), (Shape1.Top + y0))
                        SetPixel Picture3.hdc, (Shape1.Left + x0), (Shape1.Top + y0), GrdAry(x, Y)
                    End If
            Next
        Next
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0
If Check3.Value = 1 And degrees < Val(Text4.Text) Then degrees = degrees + Val(Text5.Text): GoTo redo
    Picture3.Refresh
    ShowSelection Shape1.Left, Shape1.Top
End Sub

Private Sub pastFlipHor_Click() 'switch the left pixels to the right
Dim a As Long
Dim B As Long
On Error Resume Next
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
        For a = 0 To Shape1.Width
            For B = 0 To Shape1.Height
                If GrdAry(a, B) <> 0 Then
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + (Shape1.Width - a), (Shape1.Top + B), Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, Shape1.Left + (Shape1.Width - a), (Shape1.Top + B))
                    SetPixel Picture3.hdc, Shape1.Left + (Shape1.Width - a), (Shape1.Top + B), GrdAry(a, B)
                End If
            Next B
        Next a
    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to exit the loop
    Picture3.Refresh
    ShowSelection Shape1.Left, Shape1.Top
End Sub

Private Sub PicMain_Resize()
        If viewGridRule.Checked = True Then
            PalPic.Move 16, 16, PicMain.ScaleWidth - 16, PicMain.ScaleHeight - 16
        Else
            PalPic.Move 0, 0, PicMain.ScaleWidth, PicMain.ScaleHeight
        End If
    getGridSize Val(Text2.Text) 'when the edit is resized the grid must also be resized
    ShowSelection Shape1.Left, Shape1.Top 'arter resizing the grid you will need to redisplay the image in it
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'select a color from the pallet
On Error Resume Next
    selColor = GetPixel(Picture2.hdc, x, Y)
    editPic.BackColor = GetPixel(Picture2.hdc, x, Y)
        If paternBrush = True Then
            paternBrush = False
            ReDim brush(0)
        End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Dim a As Long
Dim B As Long
Dim StepVal As Single
'block errors by reseting the x and y cord within our picture this keeps our selection shape inside the picture
        If x > Picture3.ScaleWidth - Shape1.Width Then
            x = Picture3.ScaleWidth - Shape1.Width
        ElseIf x < 0 Then
            x = 0
        End If
        If Y > Picture3.ScaleHeight - Shape1.Height Then
            Y = Picture3.ScaleHeight - Shape1.Height
        ElseIf Y < 0 Then
            Y = 0
        End If
'position the selection shape
    Shape1.Left = x
    Shape1.Top = Y
'poition the picture to keep the selection visible
        If -Picture3.Left + Picture5.ScaleWidth < Shape1.Left + Shape1.Width Then
            Picture3.Left = Picture3.Left - ((Shape1.Left + Shape1.Width) - (-Picture3.Left + Picture5.ScaleWidth))
            HScroll1.Value = -Picture3.Left
        ElseIf Shape1.Left < -Picture3.Left Then
            Picture3.Left = Picture3.Left + 1
                If Not -Picture3.Left < 0 Then HScroll1.Value = -Picture3.Left
        End If
        If -Picture3.Top + Picture5.ScaleHeight < Shape1.Top + Shape1.Height Then
            Picture3.Top = Picture3.Top - ((Shape1.Top + Shape1.Height) - (-Picture3.Top + Picture5.ScaleHeight))
            VScroll1.Value = -Picture3.Top
        ElseIf Shape1.Top < -Picture3.Top Then
            Picture3.Top = Picture3.Top + 1
                If Not -Picture3.Top < 0 Then VScroll1.Value = -Picture3.Top
        End If
    ShowSelection x, Y
        If Button = 1 Then
            If SelEnabl = True Then 'if you are selecting pixels then clear the current selection to prepare for your new selections
                ReDim SelAry(Shape1.Left To Val(Text2.Text) + Shape1.Left, Shape1.Top To Val(Text2.Text) + Shape1.Top)
                ReDim GrdAry(Val(Text2.Text), Val(Text2.Text))
                orgSel = True
            End If
        ElseIf Button = 2 Then
            oldX = x: oldY = Y
                If SelEnabl = True Then
                        For a = 0 To UBound(GrdAry, 1) - 1
                            For B = 0 To UBound(GrdAry, 2) - 1
                                If GrdAry(a, B) <> 0 Then
                                    orgSel = False
                                    Exit Sub
                                End If
                            Next B
                        Next a
                    ReDim SelAry(Shape1.Left To Val(Text2.Text) + Shape1.Left, Shape1.Top To Val(Text2.Text) + Shape1.Top)
                    ReDim GrdAry(Val(Text2.Text), Val(Text2.Text))
                    orgSel = True
'getting an error here if user tries to select in the grid after right clicking in a new area when enable select tool is enabled.
                End If
        End If
End Sub

Public Function getGridSize(vl As Long) As Long 'this must be called before any drawing to the grid image.
        If vl > Picture3.ScaleWidth Then vl = Picture3.ScaleWidth - 4 'just to ensure you don't have a selection larger than the source
        If viewGridRule.Checked = True Then 'reposition the edit grid
            PalPic.Move 16, 16, PicMain.ScaleWidth - 16, PicMain.ScaleHeight - 16
        Else
            PalPic.Move 0, 0, PicMain.ScaleWidth, PicMain.ScaleHeight
        End If
        If vl <= 0 Then Text2.Text = Shape1.Width: Exit Function 'if we enter a number bellow 1
        If vl > PalPic.ScaleWidth \ 2 Then vl = PalPic.ScaleWidth \ 2 'the max grid size can never be more than half the width
'mW is the key to the grid sizes and to the show selection sub
    mW = PalPic.ScaleWidth \ vl ' the value you pass is devided into the width or the picture box
'If your picture is 3299 you devide that by your value in this case 100 your result would be 32
'but if you split the picture up based on text2 intervals you would have 99 pixels at the end uncolored
    Text2.Text = PalPic.ScaleWidth \ mW 'therefore our value is updated to incorperate these last whole pixels
'but there would still be 3 pixels here unacounted for
    PalPic.Width = mW * Val(Text2.Text) 'so we resize the picture to exclude them
    PalPic.Height = mW * Val(Text2.Text) 'set the height to match
    Shape1.Width = Val(Text2.Text) 'resize the shape(Selection Square) based on our new size
    Shape1.Height = Val(Text2.Text)
'testing a new user control the bellow lines control rulers I created above and left of the grid
'the numbers are equal to the position in the source where the pixel is located.
        If viewGridRule = True Then
            rule1.Visible = True: rule2.Visible = True
            rule1.Move 16, 0, PalPic.Width, 15
            rule1.SmallInterval = mW
            rule1.LargeInterval = mW * 5
            rule1.NumberInterval = mW * 10
            rule2.Move 0, 16, 15, PalPic.Height
            rule2.SmallInterval = mW
            rule2.LargeInterval = mW * 5
            rule2.NumberInterval = mW * 10
        Else
            rule1.Visible = False: rule2.Visible = False
        End If
'*************************************
End Function

Private Sub DrawGrid(NumOfRow As Integer, NumOfCol As Integer, ByRef GridContainer As PictureBox)
'this sub will draw a grid onto you edit picture(palpic). This is now optional and can be turned on or off
'from the file menu.
On Error Resume Next
Dim Y As Long
Dim x As Long
    PalPic.ForeColor = gridCol
    If Val(Text2.Text) <> NumOfRow Then Text2.Text = NumOfRow
        For Y = 0 To GridContainer.ScaleHeight Step mW
            GridContainer.Line (1, Y)-(GridContainer.ScaleWidth, Y)
        Next Y
        For x = 0 To GridContainer.ScaleWidth Step mW
            GridContainer.Line (x, 1)-(x, GridContainer.ScaleHeight)
        Next x
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If ViewCord.Checked = True Then
        Picture3.ToolTipText = x & "," & Y
    Else
        Picture3.ToolTipText = ""
    End If
Dim a As Long
Dim B As Long
Dim StepVal As Single
    If Button = 1 Then
        'safty checks to ensure your selection isn't outside the picture
            If x > Picture3.ScaleWidth - Shape1.Width Then 'if we are to wide
                x = Picture3.ScaleWidth - Shape1.Width 'adjust x to position the shape inside the picture
            ElseIf x < 0 Then 'if we are left of the picture
                x = 0 'adjust x to position the shape at 0
            End If
            If Y > Picture3.ScaleHeight - Shape1.Height Then 'do the same for the top and bottom
                Y = Picture3.ScaleHeight - Shape1.Height
            ElseIf Y < 0 Then
                Y = 0
            End If
        Shape1.Left = x 'set the shapes position
        Shape1.Top = Y
        'if the shape is outside the viewable area adjust the position of the picture so yoy can see the entire selection
            If -Picture3.Left + Picture5.ScaleWidth < Shape1.Left + Shape1.Width Then
                Picture3.Left = Picture3.Left - ((Shape1.Left + Shape1.Width) - (-Picture3.Left + Picture5.ScaleWidth))
                HScroll1.Value = -Picture3.Left
            ElseIf Shape1.Left < -Picture3.Left Then
                Picture3.Left = -Shape1.Left 'Picture3.Left + 1
                    If Not -Picture3.Left < 0 Then HScroll1.Value = -Picture3.Left
            End If
            If -Picture3.Top + Picture5.ScaleHeight < Shape1.Top + Shape1.Height Then
                Picture3.Top = Picture3.Top - ((Shape1.Top + Shape1.Height) - (-Picture3.Top + Picture5.ScaleHeight))
                VScroll1.Value = -Picture3.Top
            ElseIf Shape1.Top < -Picture3.Top Then
                Picture3.Top = -Shape1.Top 'Picture3.Top + 1
                    If Not -Picture3.Top < 0 Then VScroll1.Value = -Picture3.Top
            End If
        StepVal = mW
        ShowSelection Shape1.Left, Shape1.Top
    ElseIf Button = 2 Then
        If x <> oldX Or Y <> oldY Then
            If makSel = True Then DrawFocusRect Picture3.hdc, R
        makSel = True
            If Y > oldY And x > oldX Then
                R.Bottom = oldY + (x - oldX)
                R.Left = oldX
                R.Right = x
                R.Top = oldY
            ElseIf Y < oldY And x > oldX Then
                R.Bottom = Y + (x - oldX)
                R.Left = oldX
                R.Right = x
                R.Top = Y
            ElseIf Y > oldY And x < oldX Then
                R.Bottom = oldY + (oldX - x)
                R.Left = x
                R.Right = oldX
                R.Top = oldY
            ElseIf Y < oldY And x < oldX Then
                R.Bottom = Y + (oldX - x)
                R.Left = x
                R.Right = oldX
                R.Top = Y
            End If
        DrawFocusRect Picture3.hdc, R
        Picture3.Refresh
        End If
    End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If makSel = True Then
        DrawFocusRect Picture3.hdc, R
        makSel = False
        Picture3.Refresh
        Shape1.Left = R.Left
        Shape1.Top = R.Top
        Text2.Text = R.Right - R.Left
        getGridSize Val(Text2.Text)
            If SelEnabl = True Then
                ReDim SelAry(Shape1.Left To Val(Text2.Text) + Shape1.Left, Shape1.Top To Val(Text2.Text) + Shape1.Top)
                ReDim GrdAry(Val(Text2.Text), Val(Text2.Text))
                orgSel = True
            End If
        ShowSelection Shape1.Left, Shape1.Top
        R.Bottom = -1
        R.Left = -1
        R.Right = -1
        R.Top = -1
    End If
End Sub

Private Sub Picture3_Resize()
On Error Resume Next
Dim DifHor As Long
Dim DifVer As Long
    DifHor = (Picture3.Width - Picture5.ScaleWidth)
        If DifHor > 0 Then
            HScroll1.Max = DifHor + 2
            HScroll1.LargeChange = (Picture5.ScaleWidth)
            HScroll1.SmallChange = HScroll1.LargeChange \ 4
        Else
            HScroll1.Max = 0
        End If
    DifVer = (Picture3.Height - Picture5.ScaleHeight)
        If DifVer > 0 Then
            VScroll1.Max = DifVer + 2
            VScroll1.LargeChange = (Picture5.ScaleHeight)
            VScroll1.SmallChange = VScroll1.LargeChange \ 4
        Else
            VScroll1.Max = 0
        End If
End Sub

Private Sub Picture4_Resize()
    VScroll1.Move Picture4.Width - (VScroll1.Width + 80), 0, VScroll1.Width, Picture4.Height - (HScroll1.Height + 80)
    HScroll1.Move 0, Picture4.Height - (HScroll1.Height + 80), Picture4.Width - (VScroll1.Width + 80), HScroll1.Height
    Command4(1).Move Picture4.Width - (VScroll1.Width + 80), Picture4.Height - (HScroll1.Height + 80), VScroll1.Width, HScroll1.Height
    Picture5.Move 0, 0, Picture4.Width - 365, Picture4.Height - 365
    Picture3_Resize
End Sub

Private Sub rule1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
Dim a As Long
'Debug.Print GetKeyState(VK_SHIFT)
If Button = 1 Then
    If SelEnabl = True Then
        For a = 0 To Shape1.Height
            If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                High_light_Pixel CLng(number), a, True
                SelAry(CLng(number) + Shape1.Left, a + Shape1.Top) = 0
                GrdAry(CLng(number), a) = 0
            Else
                High_light_Pixel CLng(number), a
                SelAry(CLng(number) + Shape1.Left, a + Shape1.Top) = GetPixel(Picture3.hdc, CLng(number) + Shape1.Left, a + Shape1.Top)
                GrdAry(CLng(number), a) = GetPixel(Picture3.hdc, CLng(number) + Shape1.Left, a + Shape1.Top)
            End If
        Next a
    Else
        If RuleX <> number Then
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
                For a = 0 To Shape1.Height
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + CLng(number), Shape1.Top + a, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, CLng(number) + Shape1.Left, a + Shape1.Top) 'key to start a loop of undo events
                    SetPixel Picture3.hdc, Shape1.Left + CLng(number), Shape1.Top + a, selColor
                    Fill_Pixel CLng(number), a
                Next a
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
            PalPic.Refresh: Picture3.Refresh
            ShowSelection Shape1.Left, Shape1.Top
            RuleX = number
        End If
    End If
End If
End Sub

Private Sub rule2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
If Button = 1 Then
Dim a As Long
    If SelEnabl = True Then
        For a = 0 To Shape1.Width
            If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                High_light_Pixel a, CLng(number), True
                SelAry(a + Shape1.Left, CLng(number) + Shape1.Top) = 0
                GrdAry(a, CLng(number)) = 0
            Else
                High_light_Pixel a, CLng(number)
                SelAry(a + Shape1.Left, CLng(number) + Shape1.Top) = GetPixel(Picture3.hdc, a + Shape1.Left, CLng(number) + Shape1.Top)
                GrdAry(a, CLng(number)) = GetPixel(Picture3.hdc, a + Shape1.Left, CLng(number) + Shape1.Top)
            End If
        Next a
    Else
        If RuleY <> number Then
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
                For a = 0 To Shape1.Height
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + a, CLng(number) + Shape1.Top, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, Shape1.Left + a, CLng(number) + Shape1.Top) 'key to start a loop of undo events
                    SetPixel Picture3.hdc, Shape1.Left + a, CLng(number) + Shape1.Top, selColor
                    Fill_Pixel a, CLng(number)
                Next a
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
            PalPic.Refresh: Picture3.Refresh
            ShowSelection Shape1.Left, Shape1.Top
            RuleY = number
        End If
    End If
End If
End Sub

Private Sub rule2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
Dim a As Long
    If SelEnabl = True Then
        For a = 0 To Shape1.Width
            If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                High_light_Pixel a, CLng(number), True
                SelAry(a + Shape1.Left, CLng(number) + Shape1.Top) = 0
                GrdAry(a, CLng(number)) = 0
            Else
                High_light_Pixel a, CLng(number)
                SelAry(a + Shape1.Left, CLng(number) + Shape1.Top) = GetPixel(Picture3.hdc, a + Shape1.Left, CLng(number) + Shape1.Top)
                GrdAry(a, CLng(number)) = GetPixel(Picture3.hdc, a + Shape1.Left, CLng(number) + Shape1.Top)
            End If
        Next a
    Else
        If Not RuleY = number Then
            RuleY = number
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
                For a = 0 To Shape1.Height
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + a, CLng(number) + Shape1.Top, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, Shape1.Left + a, CLng(number) + Shape1.Top) 'key to start a loop of undo events
                    SetPixel Picture3.hdc, Shape1.Left + a, CLng(number) + Shape1.Top, selColor
                    Fill_Pixel a, CLng(number)
                Next a
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
            PalPic.Refresh: Picture3.Refresh
            ShowSelection Shape1.Left, Shape1.Top
        End If
    End If
End Sub

Private Sub rule1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
Dim a As Long
    If SelEnabl = True Then
        For a = 0 To Shape1.Height - 1
            If GetKeyState(VK_SHIFT) = -127 Or GetKeyState(VK_SHIFT) = -128 Then
                High_light_Pixel CLng(number), a, True
                SelAry(CLng(number) + Shape1.Left, a + Shape1.Top) = 0
                GrdAry(CLng(number), a) = 0
                PalPic.Refresh
            Else
                High_light_Pixel CLng(number), a
                SelAry(CLng(number) + Shape1.Left, a + Shape1.Top) = GetPixel(Picture3.hdc, CLng(number) + Shape1.Left, a + Shape1.Top)
                GrdAry(CLng(number), a) = GetPixel(Picture3.hdc, CLng(number) + Shape1.Left, a + Shape1.Top)
            End If
        Next a
    Else
        If Not RuleX = number Then
            RuleX = number
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
                For a = 0 To Shape1.Height
                    UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, Shape1.Left + CLng(number), Shape1.Top + a, Picture3.Left, Picture3.Top, GetPixel(Picture3.hdc, CLng(number) + Shape1.Left, a + Shape1.Top) 'key to start a loop of undo events
                    SetPixel Picture3.hdc, Shape1.Left + CLng(number), Shape1.Top + a, selColor
                    Fill_Pixel CLng(number), a
                Next a
            UnDoStk.push Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, -1, -1, Picture3.Left, Picture3.Top, 0 'key to start a loop of undo events
            PalPic.Refresh: Picture3.Refresh
            ShowSelection Shape1.Left, Shape1.Top
        End If
    End If
End Sub

Private Sub Save_Click() 'save changes to the original picture
    Picture3.Picture = Picture3.Image
        With CD
            '.FileName = ""
            .Filter = ".bmp"
            .DefaultExt = "bmp"
            .ShowSave
        End With
        If CD.FileName <> "" Then
            SavePicture Picture3.Picture, CD.FileName
        End If
End Sub

Private Sub OpenBrush_Click()
On Error Resume Next
Dim a As Long
Dim B As Long
Dim f As Long: f = FreeFile
Dim data As String
Dim parts() As String
Dim parts2() As String
    CD.FileName = ""
    CD.DefaultExt = "bru"
    CD.Filter = "Pattern Brush Files(*.bru)|*.bru"
    CD.InitDir = App.Path & "\Resources\BurshFiles\"
    CD.ShowOpen
        If CD.FileName <> "" Then
            Open CD.FileName For Input As #f
                data = Input(LOF(f), f)
            Close #f
        End If
     parts = Split(data, vbCrLf)
     parts2 = Split(parts(0), ",")
     ReDim brush(parts2(0) To parts2(1), parts2(2) To parts2(3))
     parts2 = Split(parts(1), ",")
        For a = LBound(brush, 1) To UBound(brush, 1) - 1
            For B = LBound(brush, 2) To UBound(brush, 2) - 1
                brush(a, B) = CLng(parts2((a - LBound(brush, 1)) * (UBound(brush, 1) - LBound(brush, 1)) + (B - LBound(brush, 2))))
            Next B
        Next a
    paternBrush = True
End Sub

Private Sub SaveBrush_Click()
'check to ensure there is a brush
On Error Resume Next
Dim a As Long
Dim B As Long
Dim f As Long: f = FreeFile
Dim data As String
    MkDir App.Path & "\Resources"
    MkDir App.Path & "\Resources\BurshFiles"
    data = LBound(brush, 1) & "," & UBound(brush, 1) & "," & LBound(brush, 2) & "," & UBound(brush, 2) & vbCrLf
        For a = LBound(brush, 1) To UBound(brush, 1) - 1
            For B = LBound(brush, 2) To UBound(brush, 2) - 1
                data = data & brush(a, B) & ","
            Next B
        Next a
    CD.DefaultExt = "bru"
    CD.Filter = "Pattern Brush Files(*.bru)|*.bru"
    CD.InitDir = App.Path & "\Resources\BurshFiles\"
    CD.ShowSave
        If CD.FileName <> "" Then
            Open CD.FileName For Output As #f
                Print #f, data
            Close #f
        End If
End Sub

Private Sub SaveGridImage_Click() 'saves the current grid display including grid lines as a bmp file
    PalPic.Picture = PalPic.Image
        With CD
            .Filter = ".bmp"
            .DefaultExt = "bmp"
            .ShowSave
        End With
        If CD.FileName <> "" Then SavePicture PalPic.Picture, CD.FileName
End Sub

Private Sub saveGridMap_Click()
    Command1_Click
End Sub

Private Sub SelectAllBut_Click() 'not very useful in a true color bitmap
'need to setup an array of colors to not select
Dim x As Long
Dim Y As Long
    If Check1.Value = 1 Then
        For x = 0 To Shape1.Width
            For Y = 0 To Shape1.Height
                If GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top) <> selColor Then
                        High_light_Pixel x, Y
                        SelAry(x + Shape1.Left, Y + Shape1.Top) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top)
                        GrdAry(x, Y) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top)
                End If
            Next Y
        Next x
    End If
End Sub

Private Sub SelectAllOf_Click() 'not very useful in a true color bitmap
'need to setup an array of colors to select
Dim x As Long
Dim Y As Long
    If Check1.Value = 1 Then
        For x = 0 To Shape1.Width
            For Y = 0 To Shape1.Height
                If GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top) = selColor Then
                    High_light_Pixel x, Y
                    SelAry(x + Shape1.Left, Y + Shape1.Top) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top)
                    GrdAry(x, Y) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top)
                End If
            Next Y
        Next x
    Else 'color the pixels
    End If
End Sub

Private Sub SelInside_Click() 'this will highlight any pixels between an two other pixels in the Y direction
'Just select a pixel in the grid move down a few pixels and select another select this command and all
'the pixels between your selections will be highlighted.
'needs work; should be able to select a shape like the letter "I" blocked, where you have pixels selected
'at the top and at the bottom but none between.
    If Not Check1.Value = 1 Then Exit Sub
Dim x As Long
Dim Y As Long
Dim C As Long
Dim start As Boolean
start = False
    For x = 0 To Shape1.Width
    For Y = 0 To Shape1.Height
        'look through this colum of pixels to see if any are selected
        If GrdAry(x, Y) <> 0 Then
            If start = False Then
                start = True
            Else
                start = False
                    For C = Y + 1 To Shape1.Height
                        If GrdAry(x, C) <> 0 Then
                            start = True
                            Exit For
                        End If
                    Next C
                If start = False Then Exit For
            End If
        End If
        
        If start = True Then
            High_light_Pixel x, Y
            SelAry(x + Shape1.Left, Y + Shape1.Top) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top)
            GrdAry(x, Y) = GetPixel(Picture3.hdc, x + Shape1.Left, Y + Shape1.Top)
        End If
    Next Y
    If start = True Then start = False
    Next x
End Sub

Private Sub selPoint_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    xc = x: yc = Y
    Label6.Caption = "POINT:" & x & "," & Y
    'Label6.Refresh
End Sub

Private Sub selPoint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    selPoint.ToolTipText = x & "," & Y
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Command5_Click
End Sub


Private Sub UndoClear_Click()
'this menus' events are traped in the undo class
End Sub

Private Sub ViewCord_Click()
    ViewCord.Checked = Not ViewCord.Checked
End Sub

Private Sub ViewCross_Click()
    ViewCross.Checked = Not ViewCross.Checked
End Sub

Private Sub viewGridRule_Click()
    viewGridRule.Checked = Not viewGridRule.Checked
    PicMain_Resize
End Sub

Private Sub VScroll1_Change()
    Picture3.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_GotFocus()
    Picture3.SetFocus
End Sub

Private Sub VScroll1_Scroll()
    Picture3.Top = -VScroll1.Value
End Sub

Private Sub ShowSelection(x As Single, Y As Single)
Dim a As Long
Dim B As Long
Static lstX As Long
Static lstY As Long
    StretchBlt PalPic.hdc, 0, 0, PalPic.ScaleWidth, PalPic.ScaleHeight, Picture3.hdc, x, Y, Val(Text2.Text), Val(Text2.Text), vbSrcCopy
        If Grid.Checked = True Then DrawGrid Val(Text2.Text), Val(Text2.Text), PalPic
        If SelEnabl = True Then HighLight_Selection
    PalPic.Refresh
        If viewGridRule.Checked = True Then
            If lstX <> x Or lstY <> Y Then
                rule1.StartAt = x: lstX = x
                rule2.StartAt = Y: lstY = Y
            End If
        End If
    oldX1 = 0: oldY1 = 0
End Sub

Public Sub High_light_Pixel(x As Long, Y As Long, Optional hOff As Boolean = False)
    On Error Resume Next
        If hOff = True Then 'unhighlight but don't change the highlight of an ajoining selection
            PalPic.ForeColor = gridCol
                If Y = 0 Or GrdAry(x, Y - 1) = 0 Then PalPic.Line ((x * mW), (Y * mW))-((x * mW) + mW, (Y * mW))
                If x = 0 Or GrdAry(x - 1, Y) = 0 Then PalPic.Line ((x * mW), (Y * mW))-((x * mW), (Y * mW) + mW)
                If Y = Shape1.Height Or GrdAry(x, Y + 1) = 0 Then PalPic.Line ((x * mW), (Y * mW) + mW)-((x * mW) + mW, (Y * mW) + mW)
                If x = Shape1.Width Or GrdAry(x + 1, Y) = 0 Then PalPic.Line ((x * mW) + mW, (Y * mW))-((x * mW) + mW, (Y * mW) + mW + 1)
            If Grid.Checked = False Then ShowSelection Shape1.Left, Shape1.Top 'this repaints the grid lines if they are hiden
        Else
            PalPic.ForeColor = seleCol
            PalPic.Line ((x * mW), (Y * mW))-((x * mW) + mW, (Y * mW)) 'top
            PalPic.Line ((x * mW), (Y * mW))-((x * mW), (Y * mW) + mW) 'left
            PalPic.Line ((x * mW), (Y * mW) + mW)-((x * mW) + mW, (Y * mW) + mW) 'bottom
            PalPic.Line ((x * mW) + mW, (Y * mW))-((x * mW) + mW, (Y * mW) + mW + 1) 'right
        End If
    PalPic.ForeColor = gridCol
End Sub

Private Sub Fill_Pixel(x As Long, Y As Long) 'x and y here are the position of the pixel within the selection square
'relavent to the left and top. this then allows us to determin which block within the grid to edit
Dim StartX As Long
Dim StartY As Long
    If Grid.Checked = True Then
        StartX = (x * mW) + 1
        StartY = (Y * mW) + 1
        StretchBlt PalPic.hdc, StartX, StartY, mW - 1, mW - 1, Picture3.hdc, x + Shape1.Left, Y + Shape1.Top, 1, 1, ScrCopy
    Else
        StartX = (x * mW)
        StartY = (Y * mW)
        StretchBlt PalPic.hdc, StartX, StartY, mW, mW, Picture3.hdc, x + Shape1.Left, Y + Shape1.Top, 1, 1, ScrCopy
    End If
End Sub



