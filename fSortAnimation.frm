VERSION 5.00
Begin VB.Form fSortAnimation 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Animated Sort Algorithms"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   612
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   908
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btExpand 
      Caption         =   "Collapse"
      Height          =   330
      Left            =   12480
      TabIndex        =   13
      Top             =   2340
      Width           =   960
   End
   Begin VB.ComboBox cbSpeed 
      Height          =   315
      ItemData        =   "fSortAnimation.frx":0000
      Left            =   10590
      List            =   "fSortAnimation.frx":0029
      Style           =   2  'Dropdown-Liste
      TabIndex        =   12
      Top             =   2355
      Width           =   1215
   End
   Begin VB.Frame frSort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   135
      TabIndex        =   15
      Top             =   0
      Width           =   13320
      Begin VB.CommandButton btBreak 
         Caption         =   "Stop"
         Height          =   285
         Left            =   12615
         TabIndex        =   27
         ToolTipText     =   "Click here to interrupt"
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtInterrupt 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   10515
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Stopping, please wait"
         Top             =   0
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Shape shpUpper 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   330
         Left            =   15
         Shape           =   2  'Oval
         Top             =   210
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Shape shpLower 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   315
         Left            =   0
         Shape           =   2  'Oval
         Top             =   1740
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Shape shpColors 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   330
         Index           =   0
         Left            =   0
         Shape           =   2  'Oval
         Top             =   975
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Frame frButtons 
      Caption         =   "Sort Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   -15
      TabIndex        =   14
      Top             =   8235
      Width           =   13680
      Begin VB.CommandButton btExit 
         Caption         =   "Exit"
         Height          =   465
         Left            =   12495
         TabIndex        =   11
         Top             =   270
         Width           =   960
      End
      Begin VB.CommandButton btShuffle 
         Caption         =   "ReShuffle"
         Height          =   465
         Left            =   1215
         TabIndex        =   1
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btQuick1 
         Caption         =   "Quick 1"
         Height          =   465
         Left            =   8970
         TabIndex        =   8
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btRandomize 
         Caption         =   "Randomize"
         Height          =   465
         Left            =   210
         TabIndex        =   0
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btBubble 
         Caption         =   "Bubble"
         Height          =   465
         Left            =   2940
         TabIndex        =   2
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btInsertion 
         Caption         =   "Insertion"
         Height          =   465
         Left            =   4950
         TabIndex        =   4
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btQuick2 
         Caption         =   "Quick 2"
         Height          =   465
         Left            =   9975
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btOptQuick 
         Caption         =   "Optimized Quick 2"
         Height          =   465
         Left            =   10980
         TabIndex        =   10
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btHeap 
         Caption         =   "Heap"
         Height          =   465
         Left            =   7965
         TabIndex        =   7
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btShell 
         Caption         =   "Shell"
         Height          =   465
         Left            =   6960
         TabIndex        =   6
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btOptInsertion 
         Caption         =   "Optimized Insertion"
         Height          =   465
         Left            =   5955
         TabIndex        =   5
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton btShaker 
         Caption         =   "Cocktail Shaker"
         Height          =   465
         Left            =   3945
         TabIndex        =   3
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F8F8F8&
      DrawWidth       =   15
      Height          =   5325
      Left            =   105
      ScaleHeight     =   266.717
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleTop        =   -20
      ScaleWidth      =   52
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2865
      Width           =   13350
   End
   Begin VB.Label lblBlocks 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   5055
      TabIndex        =   25
      Top             =   2385
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Block Moves"
      Height          =   195
      Index           =   4
      Left            =   4050
      TabIndex        =   24
      Top             =   2415
      Width           =   930
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Animation Speed"
      Height          =   195
      Index           =   3
      Left            =   9300
      TabIndex        =   23
      Top             =   2415
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Comparisons"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   2415
      Width           =   900
   End
   Begin VB.Label lblComps 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   1095
      TabIndex        =   20
      Top             =   2385
      Width           =   870
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Moves"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   19
      Top             =   2415
      Width           =   480
   End
   Begin VB.Label lblMoves 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   2835
      TabIndex        =   18
      Top             =   2385
      Width           =   870
   End
   Begin VB.Label lblRecur 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "0"
      Height          =   255
      Left            =   7890
      TabIndex        =   17
      Top             =   2385
      Width           =   870
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Recursion Level"
      Height          =   195
      Index           =   2
      Left            =   6645
      TabIndex        =   16
      Top             =   2415
      Width           =   1155
   End
End
Attribute VB_Name = "fSortAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Delay           As Currency
Private Seed            As Long
Private Const LB        As Long = 1
Private Const UB        As Long = 50
Private Colors()        As Long
Private MoveCount       As Long
Private BlockCount      As Long 'counts possible block moves
Private RippleCount     As Long
Private CompCount       As Long
Private RecurLvl        As Long
Private MaxRecurLvl     As Long
Private Break           As Boolean
Private CanInterrupt    As Boolean
Private Warp            As Boolean

Private Sub btBreak_Click()

    Break = CanInterrupt
    txtInterrupt.Visible = CanInterrupt
    QueryPerformanceFrequency Delay
    Delay = Delay / 2000

End Sub

Private Sub btBubble_Click()

  Dim i         As Long
  Dim ei        As Long
  Dim Swapped   As Boolean

    Prepare "Bubble Sort"
    ei = UB
    Do
        Swapped = False
        For i = LB + 1 To ei
            IncCompCount
            If Colors(i) < Colors(i - 1) Then
                Swap i, i - 1
                Swapped = True
                If Break Then
                    Exit For 'loop varying i
                End If
            End If
        Next i
        ei = ei - 1
    Loop While Swapped And Not Break
    Unprepare

End Sub

Private Sub btExit_Click()

    Unload Me

End Sub

Private Sub btExpand_Click()

    If btExpand.Caption = "Expand" Then
        btExpand.Caption = "Collapse"
        Height = 9660
      Else 'NOT BTEXPAND.CAPTION...
        btExpand.Caption = "Expand"
        Height = 4245
    End If
    frButtons.Top = ScaleHeight - frButtons.Height + 1

End Sub

Private Sub btHeap_Click()

  Dim i As Long

    Prepare "Heap Sort"
    For i = (UB + LB) \ 2 To LB Step -1
        If Break Then
            Exit For 'loop varying i
        End If
        Percolate i, UB
    Next i
    For i = UB - 1 To LB Step -1
        If Break Then
            Exit For 'loop varying i
        End If
        Swap LB, i + 1
        Percolate LB, i
    Next i
    Unprepare

End Sub

Private Sub btInsertion_Click()

  Dim i     As Long
  Dim j     As Long
  Dim tmp   As Long

    Prepare "Insertion Sort"
    For i = LB + 1 To UB
        tmp = Colors(i)
        Drop i
        For j = i - 1 To 1 Step -1
            IncCompCount
            If Colors(j) <= tmp Then
                Exit For 'loop varying j
            End If
            Colors(j + 1) = Colors(j)
            MoveUpper j, j + 1
        Next j
        Colors(j + 1) = tmp
        SlideLower j + 1
        If Break Then
            Exit For 'loop varying i
        End If
    Next i
    Unprepare

End Sub

Private Sub btOptInsertion_Click()

  Dim i     As Long
  Dim l     As Long
  Dim r     As Long
  Dim m     As Long
  Dim tmp   As Long

    Prepare "Optimized Insertion Sort"
    For i = LB + 1 To UB
        'we know that all elements below i are already sorted so we use a binary search
        l = 1
        r = i
        Do
            m = l + (r - l) \ 2
            IncCompCount
            If Colors(i) < Colors(m) Then
                r = m
              Else 'NOT COLORS(I)...
                l = m + 1
            End If
        Loop While l < r
        If l < i Then 'we have to shift
            tmp = Colors(i)
            Drop i
            For m = i To l + 1 Step -1
                Colors(m) = Colors(m - 1)
                MoveUpper m - 1, m
                RippleCount = RippleCount + 1
            Next m
            IncblockCount -(m <> i)
            Colors(l) = tmp
            SlideLower l
        End If
        If Break Then
            Exit For 'loop varying i
        End If
    Next i
    Unprepare

End Sub

Private Sub btOptQuick_Click()

    Prepare "Optimized Quicksort"
    OptQuicksort 1, UB
    Unprepare

End Sub

Private Sub btQuick1_Click()

    Prepare "Quicksort 1"
    Quicksort1 LB, UB
    Unprepare

End Sub

Private Sub btQuick2_Click()

    Prepare "Quicksort 2"
    Quicksort2 LB, UB
    Unprepare

End Sub

Private Sub btRandomize_Click()

    Seed = Timer
    btShuffle_Click

End Sub

Private Sub btShaker_Click()

  Dim i         As Long
  Dim ai        As Long
  Dim ei        As Long
  Dim Swapped   As Boolean

    Prepare "Cocktail Shaker Sort"
    ai = LB + 1
    ei = UB
    Do Until ai = ei Or Break
        Swapped = False
        For i = ai To ei
            IncCompCount
            If Colors(i) < Colors(i - 1) Then
                Swap i, i - 1
                Swapped = True
            End If
            If Break Then
                Exit For 'loop varying i
            End If
        Next i
        If Swapped Then
            ei = ei - 1
            Swapped = False
            For i = ei To ai Step -1
                IncCompCount
                If Colors(i) < Colors(i - 1) Then
                    Swap i - 1, i
                    Swapped = True
                End If
                If Break Then
                    Exit For 'loop varying i
                End If
            Next i
        End If
        If Swapped = False Then
            Exit Do 'loop 
        End If
        ai = ai + 1
    Loop
    Unprepare

End Sub

Private Sub btShell_Click()

  Dim i         As Long
  Dim j         As Long
  Dim Gaps      As Variant
  Dim g         As Long
  Dim CurGap    As Long
  Dim Pivot     As Long

    Prepare "Shell Sort"
    'found this gap sequence at www.iti.fh-flensburg.de/lang/algorithmen/sortieren/shell/shellen.htm
    Gaps = Array(1391376, 463792, 198768, 86961, 33936, 13776, 4592, 1968, 861, 336, 112, 48, 21, 7, 3, 1, 0)
    CurGap = Gaps(g)
    Do
        For j = LB + CurGap To UB
            Pivot = Colors(j)
            Drop j
            For i = j - CurGap To 1 Step -CurGap
                IncCompCount
                If Pivot < Colors(i) Then
                    Colors(i + CurGap) = Colors(i)
                    MoveUpper i, i + CurGap
                  Else 'NOT Pivot...
                    Exit For 'loop varying i
                End If
            Next i
            Colors(i + CurGap) = Pivot
            SlideLower i + CurGap
            If Break Then
                Exit For 'loop varying j
            End If
        Next j
        g = g + 1
        CurGap = Gaps(g)
    Loop While CurGap And Not Break
    Unprepare

End Sub

Private Sub btShuffle_Click()

  Dim i As Long
  Dim c As Long

    Prepare "Shuffle"
    Rnd (-Seed)
    For i = LB To UB
        c = Rnd * 220
        Colors(i) = c
        shpColors(i).FillColor = ConvertToColor(c)
    Next i
    Display True
    Unprepare

End Sub

Private Sub cbSpeed_Click()

    If cbSpeed.ItemData(cbSpeed.ListIndex) Then
        QueryPerformanceFrequency Delay
        Delay = Delay / cbSpeed.ItemData(cbSpeed.ListIndex)
        Warp = False
      Else 'CBSPEED.ITEMDATA(CBSPEED.LISTINDEX) = FALSE/0
        Warp = True
    End If

End Sub

Private Function ConvertToColor(ByVal Value As Long) As Long

    ConvertToColor = RGB(HUEtoRGB(Value + 255 / 3), HUEtoRGB(Value), HUEtoRGB(Value - 255 / 3))

End Function

Private Sub Display(Doit As Boolean)

  Dim i As Long

    If Doit Then
        picDisplay.Cls
        For i = LB To UB
            picDisplay.PSet (i, Colors(i)), ConvertToColor(Colors(i))
        Next i
    End If

End Sub

Private Sub Drop(Idx As Long)

  Dim i As Long

    With shpLower
        .FillColor = shpColors(Idx).FillColor
        .Left = shpColors(Idx).Left
        .Top = shpColors(Idx).Top
        .Visible = True
        shpColors(Idx).Visible = False
        If Not Warp Then
            For i = 1 To 40
                .Top = .Top + 15
                Wait
            Next i
        End If
    End With 'shpLower
    IncMoveCount

End Sub

Private Sub Form_Activate()

    btExpand_Click

End Sub

Private Sub Form_Load()

  Dim i As Long

    With picDisplay
        .ScaleLeft = -0.45
        .ScaleTop = -20
        .ScaleWidth = 52
        .ScaleHeight = 250
    End With 'PICDISPLAY
    For i = LB To UB
        Load shpColors(i)
        With shpColors(i)
            .Move shpColors(i - 1).Left + shpColors(i - 1).Width + 30
            .Visible = True
        End With 'shpColors(I)
    Next i
    cbSpeed.ListIndex = 2
    ReDim Colors(LB To UB)
    btRandomize_Click

End Sub

Private Function HUEtoRGB(ByVal Hue As Long) As Long

    Select Case Hue
      Case Is < 0
        Hue = Hue + 255
      Case Is > 255
        Hue = Hue - 255
    End Select
    Select Case Hue
      Case Is < 255 / 6
        HUEtoRGB = 65 + 6 * 125 * Hue / 255
      Case Is < 255 / 2
        HUEtoRGB = 190
      Case Is < 255 * 2 / 3
        HUEtoRGB = 65 + 6 * 125 * (255 * 2 / 3 - Hue) / 255
      Case Else
        HUEtoRGB = 65
    End Select

End Function

Private Sub IncblockCount(Optional ByVal By As Long = 1)

    BlockCount = BlockCount + By
    lblBlocks = BlockCount & " for " & RippleCount

End Sub

Private Sub IncCompCount(Optional ByVal By As Long = 1)

    CompCount = CompCount + By
    lblComps = CompCount

End Sub

Private Sub IncMoveCount(Optional ByVal By As Long = 1)

    MoveCount = MoveCount + By
    lblMoves = MoveCount

End Sub

Private Sub IncRecur(Optional ByVal By As Long = 1)

    RecurLvl = RecurLvl + By
    If RecurLvl > MaxRecurLvl Then
        MaxRecurLvl = RecurLvl
    End If
    lblRecur = RecurLvl & " / " & MaxRecurLvl

End Sub

Private Sub MoveUpper(ByVal FromIdx As Long, ByVal ToIdx As Long)

  Dim i As Long

    With shpUpper
        .Left = shpColors(FromIdx).Left
        .Top = shpColors(FromIdx).Top
        .FillColor = shpColors(FromIdx).FillColor
        .Visible = True
        shpColors(FromIdx).Visible = False
        If Not Warp Then
            If Abs(FromIdx - ToIdx) > 1 Then
                For i = 1 To 40
                    .Top = .Top - 15
                    Wait
                Next i
            End If
            Do Until .Left = shpColors(ToIdx).Left
                .Left = .Left + 15 * Sgn(shpColors(ToIdx).Left - .Left)
                Wait
            Loop
            Do Until .Top = shpColors(0).Top
                .Top = .Top + 15
                Wait
            Loop
        End If
        shpColors(ToIdx).FillColor = .FillColor
        shpColors(ToIdx).Visible = True
        .Visible = False
    End With 'shpUpper
    IncMoveCount
    Display Not Warp

End Sub

Private Sub OptQuicksort(ByVal xFrom As Long, ByVal xThru As Long)

  'uses binary insertion sort and (could use) block moves for small parts

  Dim xLeft As Long
  Dim xRite As Long
  Dim l     As Long
  Dim r     As Long
  Dim Pivot As Long 'this receives table elements

    Do While xFrom < xThru And Not Break 'we have something to sort (@ least two elements)
        Select Case xThru - xFrom

          Case 1 '2 elements only: sort by swapping
            If Colors(xFrom) > Colors(xThru) Then
                Swap xFrom, xThru
            End If
            Exit Do 'done 'loop 

          Case Is < 5 'less than 6 elements: sort by insertion
            For xLeft = xFrom + 1 To xThru
                l = xFrom
                r = xLeft
                'we know that all elements below xleft are already sorted so we use a binary search
                Do 'find insertion point by binary search
                    xRite = l + (r - l) \ 2
                    IncCompCount
                    If Colors(xLeft) < Colors(xRite) Then
                        r = xRite
                      Else 'NOT COLORS(XLEFT)...
                        l = xRite + 1
                    End If
                Loop While l < r
                If l < xLeft Then ' we have to shift
                    Pivot = Colors(xLeft)
                    Drop xLeft
                    'in a production sort the following staggered move would be replaced by a block move and
                    'the limit for insertion sort activation would have to be established empirically
                    'for example CopyWordsUp l + 4, l, (left - l)
                    For xRite = xLeft To l + 1 Step -1
                        Colors(xRite) = Colors(xRite - 1)
                        MoveUpper xRite - 1, xRite
                        RippleCount = RippleCount + 1
                    Next xRite
                    IncblockCount -(xRite <> xLeft)
                    Colors(l) = Pivot
                    SlideLower l
                End If
            Next xLeft
            Exit Do 'done 'loop 

          Case Else 'many elements: sort by quicksort
            xLeft = (xFrom + xThru) / 2
            'find the median of three for a better balancing of the "halves"
            'this improves sorting a presorted sequence substantially
            Select Case True
              Case (Colors(xFrom) < Colors(xLeft)) Xor (Colors(xThru) < Colors(xLeft))
                Swap xFrom, xLeft
                IncCompCount 2
              Case (Colors(xFrom) < Colors(xThru)) Xor (Colors(xLeft) < Colors(xThru))
                Swap xFrom, xThru
                IncCompCount 4
              Case Else
                'median is alrady in xfrom
                IncCompCount 4 'but 4 comparisons anyway
            End Select

            xLeft = xFrom
            xRite = xThru
            Pivot = Colors(xLeft) 'get pivot elem and make room
            Drop xLeft
            Do
                Do Until xRite = xLeft
                    IncCompCount
                    If Colors(xRite) < Pivot Then 'is less than pivot
                        Colors(xLeft) = Colors(xRite) 'so move it to the left
                        MoveUpper xRite, xLeft
                        If Break Then
                            xLeft = xRite
                            Exit Do 'loop 
                        End If
                        xLeft = xLeft + 1 'leave the item just moved alone for now
                        Exit Do 'loop 
                    End If
                    xRite = xRite - 1
                Loop
                Do Until xLeft = xRite
                    IncCompCount
                    If Colors(xLeft) > Pivot Then 'is greater than pivot
                        Colors(xRite) = Colors(xLeft) 'so move it to the right
                        MoveUpper xLeft, xRite
                        xRite = xRite - 1 'leave the item just moved alone for now
                        Exit Do 'loop 
                    End If
                    xLeft = xLeft + 1
                Loop
            Loop Until xLeft = xRite
            'now the indexes have met and all bigger items are to the right and all smaller items are left
            Colors(xRite) = Pivot 'insert Pivot and sort the two areas left and right of it
            SlideLower xRite
            If xLeft - xFrom < xThru - xRite Then 'smaller part 1st to reduce recursion depth
                xLeft = xFrom
                xFrom = xRite + 1
                xRite = xRite - 1
              Else 'NOT XLEFT...
                xRite = xThru
                xThru = xLeft - 1
                xLeft = xLeft + 1
            End If
            If xLeft < xRite Then 'smaller part is not empty...
                IncRecur
                OptQuicksort xLeft, xRite '...so sort it
                IncRecur -1
            End If
        End Select
    Loop

End Sub

Private Sub Percolate(ByVal Lower As Long, ByVal Upper As Long)

  Dim Leaf As Long
  Dim Node As Long
  Dim tmp  As Long

    Node = Lower
    Drop Lower
    tmp = Colors(Node)
    Do
        Leaf = Node + Node - LB + 1
        If Leaf > Upper Then
            Exit Do 'loop 
        End If
        If Leaf < Upper Then
            IncCompCount
            If Colors(Leaf) < Colors(Leaf + 1) Then
                Leaf = Leaf + 1
            End If
        End If
        IncCompCount
        If Colors(Leaf) < tmp Then
            Exit Do 'loop 
        End If
        Colors(Node) = Colors(Leaf)
        MoveUpper Leaf, Node
        Node = Leaf
    Loop
    Colors(Node) = tmp
    SlideLower Node

End Sub

Private Sub Prepare(Title As String)

    frSort.Caption = Title
    frButtons.Enabled = False
    IncMoveCount -MoveCount
    IncCompCount -CompCount
    RippleCount = 0
    IncblockCount -BlockCount
    MaxRecurLvl = 0
    IncRecur -RecurLvl
    Break = False
    CanInterrupt = True
    btBreak.Visible = True

End Sub

Private Sub Quicksort1(ByVal xFrom As Long, ByVal xThru As Long)

  'conventional quicksort

  Dim xLeft As Long
  Dim xRite As Long
  Dim Pivot As Long

    If xFrom < xThru And Not Break Then
        Pivot = Colors((xFrom + xThru) \ 2)
        IncMoveCount
        xLeft = xFrom
        xRite = xThru
        Do
            Do While Colors(xLeft) < Pivot
                IncCompCount
                xLeft = xLeft + 1
            Loop
            Do While Pivot < Colors(xRite)
                IncCompCount
                xRite = xRite - 1
            Loop
            IncCompCount 2
            If xLeft <= xRite Then
                If xLeft < xRite Then
                    Swap xLeft, xRite
                End If
                xLeft = xLeft + 1
                xRite = xRite - 1
            End If
        Loop While (xLeft <= xRite) And Not Break
        IncRecur
        Quicksort1 xFrom, xRite
        Quicksort1 xLeft, xThru
        IncRecur -1
    End If

End Sub

Private Sub Quicksort2(ByVal xFrom As Long, ByVal xThru As Long)

  'enhanced  quicksort

  Dim xLeft As Long
  Dim xRite As Long
  Dim Pivot As Long 'this receives table elements

    Do While xFrom < xThru And Not Break 'we have something to sort (@ least two elements)
        xLeft = xFrom
        xRite = xThru
        Pivot = Colors(xLeft) 'get pivot elem and make room
        Drop xLeft
        Do
            Do Until xRite = xLeft
                IncCompCount
                If Colors(xRite) < Pivot Then 'is less than pivot
                    Colors(xLeft) = Colors(xRite) 'so move it to the left
                    MoveUpper xRite, xLeft
                    If Break Then
                        xLeft = xRite
                        Exit Do 'loop 
                    End If
                    xLeft = xLeft + 1 'leave the item just moved alone for now
                    Exit Do 'loop 
                End If
                xRite = xRite - 1
            Loop
            Do Until xLeft = xRite
                IncCompCount
                If Colors(xLeft) > Pivot Then 'is greater than pivot
                    Colors(xRite) = Colors(xLeft) 'so move it to the right
                    MoveUpper xLeft, xRite
                    xRite = xRite - 1 'leave the item just moved alone for now
                    Exit Do 'loop 
                End If
                xLeft = xLeft + 1
            Loop
        Loop Until xLeft = xRite
        'now the indexes have met and all bigger items are to the right and all smaller items are left
        Colors(xRite) = Pivot 'insert Pivot and sort the two areas left and right of it
        SlideLower xRite
        If xLeft - xFrom < xThru - xRite Then 'smaller part 1st to reduce recursion depth
            xLeft = xFrom
            xFrom = xRite + 1
            xRite = xRite - 1
          Else 'NOT XLEFT...
            xRite = xThru
            xThru = xLeft - 1
            xLeft = xLeft + 1
        End If
        If xLeft < xRite Then 'smaller part is not empty...
            IncRecur
            Quicksort2 xLeft, xRite '...so sort it
            IncRecur -1
        End If
    Loop

End Sub

Private Sub SlideLower(ByVal ToIdx As Long)

  Dim i As Long

    With shpLower
        If Not Warp Then
            Do Until .Left = shpColors(ToIdx).Left
                .Left = .Left + 15 * Sgn(shpColors(ToIdx).Left - .Left)
                Wait
            Loop
            For i = 1 To 40
                .Top = .Top - 15
                Wait
            Next i
        End If
        shpColors(ToIdx).Visible = True
        shpColors(ToIdx).FillColor = .FillColor
        .Visible = False
    End With 'shpLower
    IncMoveCount
    Display Not Warp

End Sub

Private Sub Swap(ByVal a As Long, ByVal b As Long)

  Dim tmp   As Long

    tmp = Colors(a)
    Drop a
    Colors(a) = Colors(b)
    MoveUpper b, a
    Colors(b) = tmp
    SlideLower b

End Sub

Private Sub Unprepare()

    frButtons.Enabled = True
    frSort.Caption = frSort.Caption & IIf(Break, " interrupted", " done")
    CanInterrupt = False
    txtInterrupt.Visible = False
    btBreak.Visible = False
    If Break Then
        cbSpeed_Click
    End If
    Display True

End Sub

Private Sub Wait()

  Dim CurTick As Currency
  Dim NxtTick As Currency

    QueryPerformanceCounter CurTick
    CurTick = CurTick + Delay
    Do
        DoEvents
        QueryPerformanceCounter NxtTick
    Loop Until NxtTick > CurTick

End Sub

':) Ulli's VB Code Formatter V2.23.12 (2007-Apr-11 14:20)  Decl: 17  Code: 781  Total: 798 Lines
':) CommentOnly: 16 (2%)  Commented: 48 (6%)  Empty: 131 (16,4%)  Max Logic Depth: 7
