VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "MapIt"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6165
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMenu 
      Align           =   4  'Align Right
      Height          =   4935
      Left            =   4830
      ScaleHeight     =   4875
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Square"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up Level"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "&Down Level"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame frmSType 
         Caption         =   "Type"
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
         Begin VB.OptionButton optSType 
            Caption         =   "Normal"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optSType 
            Caption         =   "Start"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optSType 
            Caption         =   "Stairs"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optSType 
            Caption         =   "Other"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.TextBox txtComment 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comment"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   240
   End
   Begin VB.PictureBox picMap 
      Align           =   3  'Align Left
      Height          =   4935
      Left            =   0
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Square
    'North South East West
    N As Integer
    S As Integer
    E As Integer
    W As Integer
    X As Long
    Y As Long
    Level As Integer
    Comment As String
    SType As Integer
    Part As Direct3DRMMeshBuilder3
End Type
Dim DxMap As New clsDX73D
Dim CurPos As Integer
Dim MovePos As Integer
Dim CurLevel As Integer
Dim Map() As Square

Function AddSquare() As Integer
On Error GoTo eTrap
Dim NewPos As Integer
NewPos = UBound(Map) + 1
ReDim Preserve Map(NewPos)
Map(NewPos).N = -1
Map(NewPos).S = -1
Map(NewPos).E = -1
Map(NewPos).W = -1
Map(NewPos).Level = CurLevel
AddSquare = NewPos
Exit Function
eTrap:
    NewPos = 0
    Resume Next
End Function

Function CheckNorth() As Integer
Dim i As Integer
For i = 0 To UBound(Map)
    If Map(i).Level = Map(CurPos).Level And Map(i).X = Map(CurPos).X And Map(i).Y = Map(CurPos).Y + 10 Then
        CheckNorth = i
        DxMap.mFrO.DeleteVisual Map(i).Part
        Exit Function
    End If
Next i
CheckNorth = AddSquare()
End Function

Function CheckDown() As Integer
Dim i As Integer
For i = 0 To UBound(Map)
    If Map(i).Level = Map(CurPos).Level + 1 And Map(i).X = Map(CurPos).X And Map(i).Y = Map(CurPos).Y Then
        CheckDown = i
        DxMap.mFrO.DeleteVisual Map(i).Part
        Exit Function
    End If
Next i
CheckDown = AddSquare()
End Function
Function CheckUp() As Integer
Dim i As Integer
For i = 0 To UBound(Map)
    If Map(i).Level = Map(CurPos).Level - 1 And Map(i).X = Map(CurPos).X And Map(i).Y = Map(CurPos).Y Then
        CheckUp = i
        DxMap.mFrO.DeleteVisual Map(i).Part
        Exit Function
    End If
Next i
CheckUp = AddSquare()
End Function
Function CheckSouth() As Integer
Dim i As Integer
For i = 0 To UBound(Map)
    If Map(i).Level = Map(CurPos).Level And Map(i).X = Map(CurPos).X And Map(i).Y = Map(CurPos).Y - 10 Then
        CheckSouth = i
        DxMap.mFrO.DeleteVisual Map(i).Part
        Exit Function
    End If
Next i
CheckSouth = AddSquare()
End Function
Function CheckEast() As Integer
Dim i As Integer
For i = 0 To UBound(Map)
    If Map(i).Level = Map(CurPos).Level And Map(i).X = Map(CurPos).X + 10 And Map(i).Y = Map(CurPos).Y Then
        CheckEast = i
        DxMap.mFrO.DeleteVisual Map(i).Part
        Exit Function
    End If
Next i
CheckEast = AddSquare()
End Function
Function CheckWest() As Integer
Dim i As Integer
For i = 0 To UBound(Map)
    If Map(i).Level = Map(CurPos).Level And Map(i).X = Map(CurPos).X - 10 And Map(i).Y = Map(CurPos).Y Then
        CheckWest = i
        DxMap.mFrO.DeleteVisual Map(i).Part
        Exit Function
    End If
Next i
CheckWest = AddSquare()
End Function
Function ClearAll()
On Local Error Resume Next
Dim i As Integer
For i = 0 To UBound(Map)
    DxMap.mFrO.DeleteVisual Map(i).Part
Next i
End Function
Sub DeleteSquare()
Dim i As Integer
ClearAll
For i = CurPos To UBound(Map) - 1
    Map(i) = Map(i + 1)
Next i
ReDim Preserve Map(UBound(Map) - 1)
For i = 0 To UBound(Map)
    If Map(i).N = CurPos Then Map(i).N = -1
    If Map(i).S = CurPos Then Map(i).S = -1
    If Map(i).E = CurPos Then Map(i).E = -1
    If Map(i).W = CurPos Then Map(i).W = -1
Next i
CurPos = UBound(Map)
RenderAll
SetZoom
End Sub

Function RenderAll()
Dim i As Integer
For i = 0 To UBound(Map)
    RenderSquare Map(i), i
Next i
End Function
Private Function RenderSquare(MySq As Square, Pos As Integer)
On Local Error Resume Next
Dim f As Direct3DRMFace2
Dim Tex As Direct3DRMTexture3
Dim Surf As DirectDrawSurface4
Dim r(1) As RECT
DxMap.mFrO.DeleteVisual MySq.Part
Set MySq.Part = DxMap.mDrm.CreateMeshBuilder
'----------FrontFace----------
Set f = DxMap.mDrm.CreateFace()
f.AddVertex MySq.X - 5, MySq.Y + 5, MySq.Level * 10
f.AddVertex MySq.X + 5, MySq.Y + 5, MySq.Level * 10
f.AddVertex MySq.X + 5, MySq.Y - 5, MySq.Level * 10
f.AddVertex MySq.X - 5, MySq.Y - 5, MySq.Level * 10
MySq.Part.AddFace f
'------Backface---------------
'Set f = DxMap.mDrm.CreateFace()
'f.AddVertex MySq.X - 5, MySq.Y + 5, MySq.Level * 10
'f.AddVertex MySq.X - 5, MySq.Y - 5, MySq.Level * 10
'f.AddVertex MySq.X + 5, MySq.Y - 5, MySq.Level * 10
'f.AddVertex MySq.X + 5, MySq.Y + 5, MySq.Level * 10
'MySq.Part.AddFace f
'-------Texture Map Coordinates----------------
MySq.Part.SetTextureCoordinates 0, 0, 0
MySq.Part.SetTextureCoordinates 1, 1, 0
MySq.Part.SetTextureCoordinates 2, 1, 1
MySq.Part.SetTextureCoordinates 3, 0, 1
MySq.Part.SetName Pos
'---------Create Texture and surface-------
Set Tex = DxMap.CreateUpdateableTexture(50, 50, "")
Set Surf = Tex.GetSurface(0)
Select Case MySq.SType
    Case 0: Surf.SetFillColor RGB(145, 114, 104)
    Case 1: Surf.SetFillColor RGB(145, 114, 200)
    Case 2: Surf.SetFillColor RGB(145, 200, 104)
    Case 3: Surf.SetFillColor RGB(200, 114, 104)
End Select
Me.Font.Size = 25
Surf.SetFont Me.Font
Surf.SetForeColor vbWhite
Surf.setDrawWidth 2
Surf.DrawBox 1, 1, 50, 50
If Pos = MovePos Then Surf.DrawCircle 25, 25, 10
If MySq.Comment <> "" Then Surf.DrawText 15, 2, "!", False
'-------------Draw lines depending on surrounding squares
Surf.SetForeColor RGB(66, 99, 231)
Surf.setDrawWidth 4
If MySq.N = -1 Then Surf.DrawLine 0, 2, 50, 2
If MySq.S = -1 Then Surf.DrawLine 0, 48, 50, 48
If MySq.E = -1 Then Surf.DrawLine 48, 0, 48, 50
If MySq.W = -1 Then Surf.DrawLine 2, 0, 2, 50
'----------------Update the texture--------------
Tex.Changed D3DRMTEXTURE_CHANGEDPIXELS, 0, r()
'------------------Set the texture on the mesh--------
MySq.Part.SetTexture Tex
Set f = Nothing
Set Surf = Nothing
Set Tex = Nothing
'---------Show it on the canvas----------
If MySq.Level <> CurLevel Then MySq.Part.SetQuality D3DRMRENDER_WIREFRAME
DxMap.mFrO.AddVisual MySq.Part
    
'---------Set the options-------------
If Pos = CurPos Then
    optSType(MySq.SType).Value = True
    txtComment.Text = Map(Pos).Comment
End If
End Function




Sub SaveMap(FName As String)
Dim FF As Integer
Dim i As Integer
FF = FreeFile
Open FName & ".map" For Output As #FF
For i = 0 To UBound(Map)
    With Map(i)
        Write #FF, .Level, .X, .Y, .N, .S, .E, .W, .SType, CStr(.Comment)
    End With
Next i
Close #FF
End Sub

Sub SetZoom()
Dim B As D3DRMBOX
Dim cx As Long
Dim cy As Long
Dim MaxZ As Long
DxMap.mFrO.GetHierarchyBox B
cx = (B.Max.X + B.Min.X) / 2
cy = (B.Max.Y + B.Min.Y) / 2
MaxZ = Sqr(Abs(B.Max.X - B.Min.X) ^ 2 + Abs(B.Max.Y - B.Min.Y) ^ 2) + 5
If MaxZ > 300 Then
    MaxZ = 300
    cx = Map(CurPos).X
    cy = Map(CurPos).Y
End If
DxMap.mVpt.SetBack MaxZ * 2
DxMap.mFrO.SetPosition Nothing, -cx, -cy, MaxZ
End Sub





Private Sub cmdDelete_Click()
DeleteSquare
End Sub
Private Sub cmdDown_Click()
Dim Temp As Integer
Temp = CheckDown
Map(Temp).X = Map(CurPos).X
Map(Temp).Y = Map(CurPos).Y
Map(Temp).SType = Map(CurPos).SType
CurPos = Temp
CurLevel = CurLevel + 1
Map(CurPos).Level = Map(CurPos).Level + 1
ClearAll
RenderAll
End Sub

Private Sub cmdDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap.SetFocus
End Sub


Private Sub cmdLoad_Click()
Dim FF As Integer
Dim i As Integer
Dim Res As String
Res = InputBox("Enter map name.", "Load Map")
If Res = "" Then Exit Sub
If Dir(Res & ".map") = "" Then MsgBox "Map Not Found": Exit Sub
ClearAll
Erase Map()
CurLevel = 0
CurPos = 0
MovePos = 0
FF = FreeFile
Open Res & ".map" For Input As #FF
Do While Not EOF(FF)
    ReDim Preserve Map(i)
    With Map(i)
        Input #FF, .Level, .X, .Y, .N, .S, .E, .W, .SType, .Comment
    End With
    i = i + 1
Loop
Close #FF
RenderAll
SetZoom
End Sub

Private Sub cmdSave_Click()
Dim Res As String
Res = InputBox("Enter a map name.", "Save Map")
If Res = "" Then Exit Sub
SaveMap Res
End Sub

Private Sub cmdUp_Click()
Dim Temp As Integer
Temp = CheckUp
Map(Temp).X = Map(CurPos).X
Map(Temp).Y = Map(CurPos).Y
Map(Temp).SType = Map(CurPos).SType
CurPos = Temp
CurLevel = CurLevel - 1
Map(CurPos).Level = Map(CurPos).Level - 1
ClearAll
RenderAll
End Sub

Private Sub cmdUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap.SetFocus
End Sub


Private Sub Form_Load()
DxMap.InitDx picMap
'------Add first square-------
AddSquare
Map(0).SType = 1
RenderAll
SetZoom
'----Start the refresh process------
Timer1.Interval = 1
'-------Set window topmost-------
InitWindow Me ' make it topmost
DxMap.Resize picMap ' we DID resize . . .but didn't raise an event
picMap.SetFocus
End Sub


Private Sub Form_Resize()
picMap.Width = Me.Width - 1350
End Sub


Private Sub optSType_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Map(CurPos).SType = Index

End Sub


Private Sub optSType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RenderSquare Map(CurPos), CurPos
picMap.SetFocus

End Sub


Private Sub picMap_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Temp As Integer
MovePos = CurPos
Select Case KeyCode
    Case vbKeyD
        Temp = CheckDown
        Map(CurPos).SType = 2
        Map(Temp).X = Map(CurPos).X
        Map(Temp).Y = Map(CurPos).Y
        Map(Temp).SType = Map(CurPos).SType
        CurPos = Temp
        CurLevel = CurLevel + 1
        Map(CurPos).Level = Map(CurPos).Level + 1
        ClearAll
        RenderAll
    Case vbKeyU
        Temp = CheckUp
        Map(CurPos).SType = 2
        Map(Temp).X = Map(CurPos).X
        Map(Temp).Y = Map(CurPos).Y
        Map(Temp).SType = Map(CurPos).SType
        CurPos = Temp
        CurLevel = CurLevel - 1
        Map(CurPos).Level = Map(CurPos).Level - 1
        ClearAll
        RenderAll
    Case vbKeyDelete
        DeleteSquare
        MovePos = CurPos
    Case vbKeyUp
        MovePos = CheckNorth()
        Map(CurPos).N = MovePos
        Map(MovePos).S = CurPos
        Map(MovePos).X = Map(CurPos).X
        Map(MovePos).Y = Map(CurPos).Y + 10
    Case vbKeyDown
        MovePos = CheckSouth()
        Map(CurPos).S = MovePos
        Map(MovePos).N = CurPos
        Map(MovePos).X = Map(CurPos).X
        Map(MovePos).Y = Map(CurPos).Y - 10
    Case vbKeyLeft
        MovePos = CheckWest()
        Map(CurPos).W = MovePos
        Map(MovePos).E = CurPos
        Map(MovePos).X = Map(CurPos).X - 10
        Map(MovePos).Y = Map(CurPos).Y
    Case vbKeyRight
        MovePos = CheckEast()
        Map(CurPos).E = MovePos
        Map(MovePos).W = CurPos
        Map(MovePos).X = Map(CurPos).X + 10
        Map(MovePos).Y = Map(CurPos).Y
End Select
'--------Update the canvas with the new objects
RenderSquare Map(CurPos), CurPos
CurPos = MovePos
RenderSquare Map(CurPos), MovePos
SetZoom
SaveMap "TEMP"
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Pick As Integer
    DxMap.MouseDown X, Y
    Pick = -1
    If DxMap.Pick(X, Y) <> "" Then Pick = CInt(DxMap.Pick(X, Y))
    If Pick <> -1 Then
        MovePos = Pick
        CurPos = Pick
        CurLevel = Map(Pick).Level
        ClearAll
        RenderAll
    End If
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DxMap.MouseMove X, Y
End Sub


Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DxMap.MouseUp
End Sub

Private Sub picMap_Resize()
DxMap.Resize picMap
End Sub

Private Sub Timer1_Timer()
DxMap.Update
End Sub


Private Sub txtComment_Change()
Map(CurPos).Comment = txtComment.Text
End Sub


