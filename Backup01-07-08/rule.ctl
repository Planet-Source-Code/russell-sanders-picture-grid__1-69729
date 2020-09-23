VERSION 5.00
Begin VB.UserControl rule 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
End
Attribute VB_Name = "rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private SmallLength As Long
Private LargeLength As Long
Private NumberLength As Long
Private mSmallInterval As Long
Const m_def_SmallInterval As Long = 2
Private mLargeInterval As Long
Const m_def_LargeInterval As Long = 10
Private mNumberInterval As Long
Const m_def_NumberInterval As Long = 50

Public Enum typCount
    WholeNumbers 'only counts each numberinterval
    Incremental 'counts all increments
End Enum
Public Enum myBorderStyle
    None
    Fixed
End Enum
Public Enum mAppearance 'Appearance
    Flat
    Three_D
End Enum
Public Enum setOrentation
    HorLeftRight 'top or bottom numbers go left to right
    HorRightLeft 'top or bottom numbers go Right to Left
    VerTopBot 'left or right numbers go top to bottom
    VerBotTop 'left or right numbers go bottom to top
    'Drawing on both sides of the control
    HorLeftRightBoth
    HorRightLeftBoth
    VerTopBotBoth
    VerBotTopBoth
End Enum
Const m_def_Orentation = 0
Dim m_Orentation As setOrentation

Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single, number As String)
Private oldX As Integer
Private oldY As Integer
Private mStartAt As Long
Const m_def_StartAt As Long = 0

Private Type Size
    cx As Long
    cy As Long
End Type

Private TextSz As Size
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private minVertit As Boolean 'flip the drawing rutine
Const m_def_inVertit As Boolean = False
Private mCountType As Long
Const m_def_CountType As Long = 0
Private mSmallInterval2 As Long
Const m_def_SmallInterval2 As Long = 25
Private mLargeInterval2 As Long
Const m_def_LargeInterval2 As Long = 50
Private mNumberInterval2 As Long
Const m_def_NumberInterval2 As Long = 100

Public Property Get StartAt() As Long
    StartAt = mStartAt
End Property
Public Property Let StartAt(NewValue As Long)
    mStartAt = NewValue
    PropertyChanged StartAt
    Draw
End Property

Private Sub UserControl_Initialize()
    oldX = 0: oldY = 0
    PointOn 0, 0, vbRed
End Sub

Private Sub UserControl_Resize()
    oldX = 0: oldY = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SmallInterval", mSmallInterval, m_def_SmallInterval)
    Call PropBag.WriteProperty("NumberInterval2", mNumberInterval2, m_def_NumberInterval2)
    Call PropBag.WriteProperty("LargeInterval2", mLargeInterval2, m_def_LargeInterval2)
    Call PropBag.WriteProperty("SmallInterval2", mSmallInterval2, m_def_SmallInterval2)
    Call PropBag.WriteProperty("CountType", mCountType, m_def_CountType)
    Call PropBag.WriteProperty("inVertit", minVertit, m_def_inVertit)
    Call PropBag.WriteProperty("StartAt", mStartAt, m_def_StartAt)
    Call PropBag.WriteProperty("NumberInterval", mNumberInterval, m_def_NumberInterval)
    Call PropBag.WriteProperty("LargeInterval", mLargeInterval, m_def_LargeInterval)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Orentation", m_Orentation, m_def_Orentation)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 0)
    mNumberInterval2 = PropBag.ReadProperty("NumberInterval2", 100)
    mLargeInterval2 = PropBag.ReadProperty("LargeInterval2", 50)
    mSmallInterval2 = PropBag.ReadProperty("SmallInterval2", 25)
    mCountType = PropBag.ReadProperty("CountType", m_def_CountType)
    mSmallInterval = PropBag.ReadProperty("SmallInterval", 2)
    minVertit = PropBag.ReadProperty("inVertit", m_def_inVertit)
    mStartAt = PropBag.ReadProperty("StartAt", 0)
    mNumberInterval = PropBag.ReadProperty("NumberInterval", 50)
    mLargeInterval = PropBag.ReadProperty("LargeInterval", 10)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Orentation = PropBag.ReadProperty("Orentation", 0)
End Sub

Public Sub Draw()
On Error Resume Next
Dim a As Long, CurNum As Long, B As Long, CurNum2 As Long
Dim cter As Long
CurNum = StartAt
CurNum2 = StartAt
oldX = 0: oldY = 0
    UserControl.Cls
        Select Case m_Orentation
            Case 0, 1 'drawing on one side
                SmallLength = UserControl.ScaleHeight * 0.25
                LargeLength = UserControl.ScaleHeight * 0.4
                NumberLength = UserControl.ScaleHeight * 0.75
            Case 2, 3
                SmallLength = UserControl.ScaleWidth * 0.25
                LargeLength = UserControl.ScaleWidth * 0.4
                NumberLength = UserControl.ScaleWidth * 0.75
            Case 4, 5 'drawing on both sides
                SmallLength = UserControl.ScaleHeight * 0.125
                LargeLength = UserControl.ScaleHeight * 0.2
                NumberLength = UserControl.ScaleHeight * 0.375
            Case 6, 7
                SmallLength = UserControl.ScaleWidth * 0.125
                LargeLength = UserControl.ScaleWidth * 0.2
                NumberLength = UserControl.ScaleWidth * 0.375
        End Select
Select Case m_Orentation
    Case 0
            For a = 0 To UserControl.ScaleWidth
                If a Mod mSmallInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (a, UserControl.ScaleHeight - SmallLength)-(a, UserControl.ScaleHeight)
                    Else
                        UserControl.Line (a, 0)-(a, SmallLength)
                    End If
                End If
                If a Mod mLargeInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (a, UserControl.ScaleHeight - LargeLength)-(a, UserControl.ScaleHeight)
                    Else
                        UserControl.Line (a, 0)-(a, LargeLength)
                    End If
                End If
                If a Mod mNumberInterval = 0 Then
                        If Not a < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                        If inVertit = True Then
                            UserControl.Line (a, UserControl.ScaleHeight - NumberLength)-(a, UserControl.ScaleHeight)
                        Else
                            UserControl.Line (a, 0)-(a, NumberLength)
                        End If
                    UserControl.CurrentX = a
                    GetTextExtentPoint32 UserControl.hdc, CurNum, 1, TextSz
                    If inVertit = True Then
                            UserControl.CurrentY = 1
                        Else
                            UserControl.CurrentY = UserControl.ScaleHeight - (TextSz.cy - 2)
                        End If
                    UserControl.Print CurNum
                End If
            Next a
    Case 1
        cter = 0
            For a = UserControl.ScaleWidth To 0 Step -1
                If cter Mod mSmallInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (a, UserControl.ScaleHeight - SmallLength)-(a, UserControl.ScaleHeight)
                    Else
                        UserControl.Line (a, 0)-(a, SmallLength)
                    End If
                End If
                If cter Mod mLargeInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (a, UserControl.ScaleHeight - LargeLength)-(a, UserControl.ScaleHeight)
                    Else
                        UserControl.Line (a, 0)-(a, LargeLength)
                    End If
                End If
                If cter Mod mNumberInterval = 0 Then
                        If Not cter < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                        If inVertit = True Then
                            UserControl.Line (a, UserControl.ScaleHeight - NumberLength)-(a, UserControl.ScaleHeight)
                        Else
                            UserControl.Line (a, 0)-(a, NumberLength)
                        End If
                    GetTextExtentPoint32 UserControl.hdc, CurNum, Len(CStr(CurNum)) + 1, TextSz
                    UserControl.CurrentX = a - (TextSz.cx)
                        If inVertit = True Then
                            UserControl.CurrentY = 1
                        Else
                            UserControl.CurrentY = UserControl.ScaleHeight - (TextSz.cy - 2)
                        End If
                    UserControl.Print CurNum
                    
                End If
            cter = cter + 1
            Next a
    Case 2
            For a = 0 To UserControl.ScaleHeight
                If a Mod SmallInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (UserControl.ScaleWidth - SmallLength, a)-(UserControl.ScaleWidth, a)
                    Else
                        UserControl.Line (0, a)-(SmallLength, a)
                    End If
                End If
                If a Mod LargeInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (UserControl.ScaleWidth - LargeLength, a)-(UserControl.ScaleWidth, a)
                    Else
                        UserControl.Line (0, a)-(LargeLength, a)
                    End If
                End If
                If a Mod NumberInterval = 0 Then
                        If Not a < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                        If inVertit = True Then
                            UserControl.Line (UserControl.ScaleWidth - NumberLength, a)-(UserControl.ScaleWidth, a)
                        Else
                            UserControl.Line (0, a)-(NumberLength, a)
                        End If
                    UserControl.CurrentX = UserControl.ScaleWidth \ 4
                    UserControl.CurrentY = a
                    GetTextExtentPoint32 UserControl.hdc, CurNum, 1, TextSz
                        For B = 0 To Len(CurNum) '- 1
                                If inVertit = True Then
                                    UserControl.CurrentX = 1
                                Else
                                    UserControl.CurrentX = UserControl.ScaleWidth - (TextSz.cx + 2)  '* 0.4
                                End If
                            UserControl.Print Mid(CurNum, B, 1)
                        Next B
                End If
            Next a
    Case 3
    cter = 0
            For a = UserControl.ScaleHeight To 0 Step -1
                If cter Mod SmallInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (UserControl.ScaleWidth - SmallLength, a)-(UserControl.ScaleWidth, a)
                    Else
                        UserControl.Line (0, a)-(SmallLength, a)
                    End If
                End If
                If cter Mod LargeInterval = 0 Then
                    If inVertit = True Then
                        UserControl.Line (UserControl.ScaleWidth - LargeLength, a)-(UserControl.ScaleWidth, a)
                    Else
                        UserControl.Line (0, a)-(LargeLength, a)
                    End If
                End If
                If cter Mod NumberInterval = 0 Then
                        If Not cter < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                        If inVertit = True Then
                            UserControl.Line (UserControl.ScaleWidth - NumberLength, a)-(UserControl.ScaleWidth, a)
                        Else
                            UserControl.Line (0, a)-(NumberLength, a)
                        End If
                    UserControl.CurrentX = UserControl.ScaleWidth \ 4
                    GetTextExtentPoint32 UserControl.hdc, CurNum, 1, TextSz
                    UserControl.CurrentY = a - (TextSz.cy * Len(CStr(CurNum)) - 1)
                        For B = 0 To Len(CurNum) '- 1
                                If inVertit = True Then
                                    UserControl.CurrentX = 1
                                Else
                                    UserControl.CurrentX = UserControl.ScaleWidth - (TextSz.cx)  '* 0.4
                                End If
                            UserControl.Print Mid(CurNum, B, 1)
                        Next B
                End If
                cter = cter + 1
            Next a
        Case 4 'Double Drawn ver. left to right
            For a = 0 To UserControl.ScaleWidth
                If a Mod mSmallInterval = 0 Then
                    UserControl.Line (a, 0)-(a, SmallLength)
                End If
                If a Mod mSmallInterval2 = 0 Then
                    UserControl.Line (a, UserControl.ScaleHeight - SmallLength)-(a, UserControl.ScaleHeight)
                End If
                If a Mod mLargeInterval = 0 Then
                    UserControl.Line (a, 0)-(a, LargeLength)
                End If
                If a Mod mLargeInterval2 = 0 Then
                    UserControl.Line (a, UserControl.ScaleHeight - LargeLength)-(a, UserControl.ScaleHeight)
                End If
                If a Mod mNumberInterval = 0 Then
                        If Not a < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                    UserControl.Line (a, 0)-(a, NumberLength)
                    UserControl.CurrentX = a
                    GetTextExtentPoint32 UserControl.hdc, CurNum, 1, TextSz
                    UserControl.CurrentY = NumberLength - (TextSz.cy - 2)  '(UserControl.ScaleHeight \ 2) - (TextSz.cy + 2)
                    UserControl.Print CurNum
                End If
                If a Mod mNumberInterval2 = 0 Then
                        If Not a < mNumberInterval2 Then
                            If CountType = 0 Then
                                CurNum2 = CurNum2 + 1
                            Else
                                CurNum2 = CurNum2 + (mNumberInterval2 \ mSmallInterval2)
                            End If
                        End If
                    UserControl.Line (a, UserControl.ScaleHeight - NumberLength)-(a, UserControl.ScaleHeight)
                    UserControl.CurrentX = a
                    GetTextExtentPoint32 UserControl.hdc, CurNum2, 1, TextSz
                    UserControl.CurrentY = (UserControl.ScaleHeight - NumberLength) - 2 ' TextSz.cy        '+ TextSz.cy
                    UserControl.Print CurNum2
                End If
            Next a
        Case 5 'double drawn Ver. Right To Left
        cter = 0
            For a = UserControl.ScaleWidth To 0 Step -1
                If cter Mod mSmallInterval = 0 Then
                    UserControl.Line (a, 0)-(a, SmallLength)
                End If
                If cter Mod mSmallInterval2 = 0 Then
                    UserControl.Line (a, UserControl.ScaleHeight - SmallLength)-(a, UserControl.ScaleHeight)
                End If
                If cter Mod mLargeInterval = 0 Then
                    UserControl.Line (a, 0)-(a, LargeLength)
                End If
                If cter Mod mLargeInterval2 = 0 Then
                    UserControl.Line (a, UserControl.ScaleHeight - LargeLength)-(a, UserControl.ScaleHeight)
                End If
                
                If cter Mod mNumberInterval = 0 Then
                        If Not cter < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                    UserControl.Line (a, 0)-(a, NumberLength)
                    GetTextExtentPoint32 UserControl.hdc, CurNum, Len(CStr(CurNum)) + 1, TextSz
                    UserControl.CurrentX = a - (TextSz.cx)
                    UserControl.CurrentY = NumberLength - (TextSz.cy - 2) '(UserControl.ScaleHeight \ 2) - TextSz.cy
                    UserControl.Print CurNum
                End If
                If cter Mod mNumberInterval2 = 0 Then
                        If Not cter < mNumberInterval2 Then
                            If CountType = 0 Then
                                CurNum2 = CurNum2 + 1
                            Else
                                CurNum2 = CurNum2 + (mNumberInterval2 \ mSmallInterval2)
                            End If
                        End If
                    UserControl.Line (a, UserControl.ScaleHeight - NumberLength)-(a, UserControl.ScaleHeight)
                    GetTextExtentPoint32 UserControl.hdc, CurNum2, Len(CStr(CurNum)) + 1, TextSz
                    UserControl.CurrentX = a - (TextSz.cx)
                    UserControl.CurrentY = (UserControl.ScaleHeight - NumberLength) - 2 'UserControl.ScaleHeight \ 2 ' - (TextSz.cy - 2)
                    UserControl.Print CurNum2
                End If
            cter = cter + 1
            Next a
        Case 6 'Double drawn ver. Top To Bot
            For a = 0 To UserControl.ScaleHeight
                If a Mod SmallInterval = 0 Then
                    UserControl.Line (0, a)-(SmallLength, a)
                End If
                If a Mod SmallInterval2 = 0 Then
                    UserControl.Line (UserControl.ScaleWidth - SmallLength, a)-(UserControl.ScaleWidth, a)
                End If
                If a Mod LargeInterval = 0 Then
                    UserControl.Line (0, a)-(LargeLength, a)
                End If
                If a Mod LargeInterval2 = 0 Then
                    UserControl.Line (UserControl.ScaleWidth - LargeLength, a)-(UserControl.ScaleWidth, a)
                End If
                If a Mod NumberInterval = 0 Then
                        If Not a < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                    UserControl.Line (0, a)-(NumberLength, a)
                    UserControl.CurrentX = UserControl.ScaleWidth \ 4
                    UserControl.CurrentY = a
                    GetTextExtentPoint32 UserControl.hdc, CurNum, 1, TextSz
                        For B = 0 To Len(CurNum) '- 1
                            UserControl.CurrentX = (UserControl.ScaleWidth \ 2) - TextSz.cx - (UserControl.ScaleWidth \ 8)
                            UserControl.Print Mid(CurNum, B, 1)
                        Next B
                End If
                
                If a Mod NumberInterval2 = 0 Then
                        If Not a < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum2 = CurNum2 + 1
                            Else
                                CurNum2 = CurNum2 + (mNumberInterval2 \ mSmallInterval2)
                            End If
                        End If
                    UserControl.Line (UserControl.ScaleWidth - NumberLength, a)-(UserControl.ScaleWidth, a)
                    UserControl.CurrentX = UserControl.ScaleWidth \ 4
                    UserControl.CurrentY = a
                    GetTextExtentPoint32 UserControl.hdc, CurNum2, 1, TextSz
                        For B = 0 To Len(CurNum2) '- 1
                            UserControl.CurrentX = (UserControl.ScaleWidth \ 2) + (UserControl.ScaleWidth \ 8)
                            UserControl.Print Mid(CurNum2, B, 1)
                        Next B
                End If
            Next a
        Case 7 'Double Drawn Ver. Bot To Top
    cter = 0
            For a = UserControl.ScaleHeight To 0 Step -1
                If cter Mod SmallInterval = 0 Then
                    UserControl.Line (0, a)-(SmallLength, a)
                End If
                If cter Mod SmallInterval2 = 0 Then
                    UserControl.Line (UserControl.ScaleWidth - SmallLength, a)-(UserControl.ScaleWidth, a)
                End If
                If cter Mod LargeInterval = 0 Then
                    UserControl.Line (0, a)-(LargeLength, a)
                End If
                If cter Mod LargeInterval2 = 0 Then
                    UserControl.Line (UserControl.ScaleWidth - LargeLength, a)-(UserControl.ScaleWidth, a)
                End If
                If cter Mod NumberInterval = 0 Then
                        If Not cter < mNumberInterval Then
                            If CountType = 0 Then
                                CurNum = CurNum + 1
                            Else
                                CurNum = CurNum + (mNumberInterval \ mSmallInterval)
                            End If
                        End If
                    UserControl.Line (0, a)-(NumberLength, a)
                    UserControl.CurrentX = UserControl.ScaleWidth \ 4
                    GetTextExtentPoint32 UserControl.hdc, CurNum, 1, TextSz
                    UserControl.CurrentY = a - (TextSz.cy * Len(CStr(CurNum)) - 1)
                        For B = 0 To Len(CurNum)
                            UserControl.CurrentX = (UserControl.ScaleWidth \ 2) - (TextSz.cx) - (UserControl.ScaleWidth \ 8)
                            UserControl.Print Mid(CurNum, B, 1)
                        Next B
                End If
                If cter Mod NumberInterval2 = 0 Then
                        If Not cter < mNumberInterval2 Then
                            If CountType = 0 Then
                                CurNum2 = CurNum2 + 1
                            Else
                                CurNum2 = CurNum2 + (mNumberInterval2 \ mSmallInterval2)
                            End If
                        End If
                    UserControl.Line (UserControl.ScaleWidth - NumberLength, a)-(UserControl.ScaleWidth, a)
                    UserControl.CurrentX = UserControl.ScaleWidth \ 4
                    GetTextExtentPoint32 UserControl.hdc, CurNum2, 1, TextSz
                    UserControl.CurrentY = a - (TextSz.cy * Len(CStr(CurNum)) - 1)
                        For B = 0 To Len(CurNum) '- 1
                            UserControl.CurrentX = (UserControl.ScaleWidth \ 2) + (UserControl.ScaleWidth \ 8)
                            UserControl.Print Mid(CurNum2, B, 1)
                        Next B
                End If
                cter = cter + 1
            Next a
    End Select
End Sub

Private Sub UserControl_Paint()
    Draw
End Sub

Public Property Get SmallInterval2() As Long
    SmallInterval2 = mSmallInterval2
End Property
Public Property Let SmallInterval2(NewValue As Long)

    mSmallInterval2 = NewValue
PropertyChanged SmallInterval2
Draw
End Property

Public Property Get LargeInterval2() As Long
    LargeInterval2 = mLargeInterval2
End Property
Public Property Let LargeInterval2(NewValue As Long)

    mLargeInterval2 = NewValue
    PropertyChanged LargeInterval2
    Draw
End Property

Public Property Get NumberInterval2() As Long
    NumberInterval2 = mNumberInterval2
End Property
Public Property Let NumberInterval2(NewValue As Long)
    mNumberInterval2 = NewValue
    PropertyChanged NumberInterval2
    Draw
End Property


Public Property Get SmallInterval() As Long
    SmallInterval = mSmallInterval
End Property
Public Property Let SmallInterval(NewValue As Long)
    mSmallInterval = NewValue
    PropertyChanged SmallInterval
    Draw
End Property

Public Property Get LargeInterval() As Long
    LargeInterval = mLargeInterval
End Property
Public Property Let LargeInterval(NewValue As Long)
    mLargeInterval = NewValue
    PropertyChanged LargeInterval
    Draw
End Property

Public Property Get NumberInterval() As Long
    NumberInterval = mNumberInterval
End Property
Public Property Let NumberInterval(NewValue As Long)
    mNumberInterval = NewValue
    PropertyChanged NumberInterval
    Draw
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Draw
End Property

Public Property Get BorderStyle() As myBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property
'
Public Property Let BorderStyle(ByVal New_BorderStyle As myBorderStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    Draw
End Property

Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Draw
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Draw
End Property

Public Function Calculate(x As Single, Y As Single) As String
Dim posit As String
    Select Case m_Orentation
        Case HorLeftRight
                If CountType = WholeNumbers Then
                    posit = (x / mNumberInterval)
                Else
                    posit = (x \ mSmallInterval)
                End If
        Case HorRightLeft
                If CountType = WholeNumbers Then
                    posit = (UserControl.ScaleWidth / mNumberInterval) - (x / mNumberInterval)
                Else
                    posit = (UserControl.ScaleWidth \ SmallInterval) - (x \ mSmallInterval)
                End If
        Case VerTopBot
            If CountType = WholeNumbers Then
                posit = (Y / mNumberInterval)
            Else
                posit = (Y \ mSmallInterval)
            End If
        Case VerBotTop
            If CountType = WholeNumbers Then
                posit = (UserControl.ScaleHeight / mNumberInterval) - (Y / mNumberInterval)
            Else
                posit = (UserControl.ScaleHeight \ SmallInterval) - (Y \ mSmallInterval)
            End If
        Case HorLeftRightBoth
            If CountType = WholeNumbers Then
                posit = Format((x / mNumberInterval), ".00")
            Else
                posit = (x \ mSmallInterval)
            End If
            If CountType = WholeNumbers Then
                posit = posit & "," & (x / mNumberInterval2)
            Else
                posit = posit & "," & (x \ mSmallInterval2)
            End If
        Case HorRightLeftBoth
            If CountType = WholeNumbers Then
                posit = Format((UserControl.ScaleWidth / mNumberInterval) - (x / mNumberInterval), ".00")
            Else
                posit = (UserControl.ScaleWidth \ SmallInterval) - (x \ mSmallInterval)
            End If
            If CountType = WholeNumbers Then
                posit = posit & "," & Format((UserControl.ScaleWidth / mNumberInterval2) - (x / mNumberInterval2), ".00")
            Else
                posit = posit & "," & (UserControl.ScaleWidth \ SmallInterval2) - (x \ mSmallInterval2)
            End If
        Case VerTopBotBoth
            If CountType = WholeNumbers Then
                posit = Format((Y / mNumberInterval), ".00")
            Else
                posit = (Y \ mSmallInterval)
            End If
            If CountType = WholeNumbers Then
                posit = posit & "," & Format((Y / mNumberInterval2), ".00")
            Else
                posit = posit & "," & (Y \ mSmallInterval2)
            End If
        Case VerBotTopBoth
            If CountType = WholeNumbers Then
                posit = Format((UserControl.ScaleHeight / mNumberInterval) - (Y / mNumberInterval), ".00")
            Else
                posit = (UserControl.ScaleHeight \ SmallInterval) - (Y \ mSmallInterval)
            End If
            If CountType = WholeNumbers Then
                posit = posit & "," & Format((UserControl.ScaleHeight / mNumberInterval2) - (Y / mNumberInterval2), ".00")
            Else
                posit = posit & "," & (UserControl.ScaleHeight \ SmallInterval2) - (Y \ mSmallInterval2)
            End If
    End Select

Calculate = posit
End Function


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y, Calculate(x, Y))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y, Calculate(x, Y))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y, Calculate(x, Y))
End Sub

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Orentation() As setOrentation
    Orentation = m_Orentation
End Property

Public Property Let Orentation(ByVal New_Orentation As setOrentation)
    m_Orentation = New_Orentation
    PropertyChanged "Orentation"
    Draw
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Orentation = m_def_Orentation
    mNumberInterval = 50
    mSmallInterval = 2
    mLargeInterval = 10
    mNumberInterval2 = 100
    mSmallInterval2 = 25
    mLargeInterval2 = 50
End Sub

Public Sub PointOn(x As Single, Y As Single, Color As Long)
    UserControl.DrawMode = 7 'XorPen
    UserControl.DrawWidth = 2
        Select Case m_Orentation
        Case 0, 4
            If oldX <> 0 Then UserControl.Line (oldX, 0)-(oldX, UserControl.ScaleHeight), Color
            UserControl.Line (x, 0)-(x, UserControl.ScaleHeight), Color
        Case 1, 5
            If oldX <> 0 Then UserControl.Line (UserControl.ScaleWidth - oldX, 0)-(UserControl.ScaleWidth - oldX, UserControl.ScaleHeight), Color
            UserControl.Line (UserControl.ScaleWidth - x, 0)-(UserControl.ScaleWidth - x, UserControl.ScaleHeight), Color
        Case 2, 6
            If oldY <> 0 Then UserControl.Line (0, oldY)-(UserControl.ScaleWidth, oldY), Color
            UserControl.Line (0, Y)-(UserControl.ScaleWidth, Y), Color
        Case 3, 7
            If oldY <> 0 Then UserControl.Line (0, UserControl.ScaleHeight - oldY)-(UserControl.ScaleWidth, UserControl.ScaleHeight - oldY), Color
            UserControl.Line (0, UserControl.ScaleHeight - Y)-(UserControl.ScaleWidth, UserControl.ScaleHeight - Y), Color
        End Select
    oldX = x: oldY = Y
    UserControl.DrawMode = 13 'copyPen
    UserControl.DrawWidth = 1
End Sub

Public Property Get inVertit() As Boolean
    inVertit = minVertit
End Property
Public Property Let inVertit(NewValue As Boolean)
    minVertit = NewValue
    PropertyChanged inVertit
    Draw
End Property

Public Property Get Appearance() As mAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As mAppearance)
Dim bc As Long
bc = UserControl.BackColor
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
    UserControl.BackColor = bc
    Draw
End Property

Public Property Get CountType() As typCount
    CountType = mCountType
End Property
Public Property Let CountType(NewValue As typCount)
    mCountType = NewValue
    PropertyChanged CountType
    Draw
End Property


