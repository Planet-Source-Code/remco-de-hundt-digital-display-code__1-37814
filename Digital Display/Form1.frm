VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   13
      Left            =   225
      Max             =   2000
      TabIndex        =   0
      Top             =   135
      Width           =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   0
      X1              =   1890
      X2              =   2100
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   1
      X1              =   1890
      X2              =   2100
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   2
      X1              =   1890
      X2              =   2100
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   3
      X1              =   1800
      X2              =   1800
      Y1              =   885
      Y2              =   705
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   4
      X1              =   2190
      X2              =   2190
      Y1              =   885
      Y2              =   705
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   5
      X1              =   1800
      X2              =   1800
      Y1              =   1185
      Y2              =   1005
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   6
      X1              =   2190
      X2              =   2190
      Y1              =   1185
      Y2              =   1005
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   0
      X1              =   1320
      X2              =   1530
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   1
      X1              =   1320
      X2              =   1530
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   2
      X1              =   1320
      X2              =   1530
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   3
      X1              =   1230
      X2              =   1230
      Y1              =   885
      Y2              =   705
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   4
      X1              =   1620
      X2              =   1620
      Y1              =   885
      Y2              =   705
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   5
      X1              =   1230
      X2              =   1230
      Y1              =   1185
      Y2              =   1005
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   6
      X1              =   1620
      X2              =   1620
      Y1              =   1185
      Y2              =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   6
      X1              =   1050
      X2              =   1050
      Y1              =   1185
      Y2              =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   5
      X1              =   660
      X2              =   660
      Y1              =   1185
      Y2              =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   4
      X1              =   1050
      X2              =   1050
      Y1              =   885
      Y2              =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   3
      X1              =   660
      X2              =   660
      Y1              =   885
      Y2              =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   2
      X1              =   750
      X2              =   960
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   1
      X1              =   750
      X2              =   960
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000060&
      BorderWidth     =   6
      Index           =   0
      X1              =   750
      X2              =   960
      Y1              =   645
      Y2              =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Digit(number As Integer)
If number > 999 Then number = 999
Dim dg As String
Dim nmb As String
Dim s As Byte

nmb = Trim(Str(number))
nmb = Space(3 - Len(nmb)) + nmb
For k = 0 To 2
    dg = Mid(nmb, k + 1, 1)
    Select Case dg
    Case 0: dg = "1011111"
    Case 1: dg = "0000101"
    Case 2: dg = "1110110"
    Case 3: dg = "1110101"
    Case 4: dg = "0101101"
    Case 5: dg = "1111001"
    Case 6: dg = "1111011"
    Case 7: dg = "1000101"
    Case 8: dg = "1111111"
    Case 9: dg = "1111101"
    Case Else: dg = "0000000"
    End Select
    
    For l = 0 To 6
        s = Mid(dg, l + 1, 1)
        Select Case k
            Case 0: If s = 0 Then Line1(l).BorderColor = &H60& Else Line1(l).BorderColor = &HFF&
            Case 1: If s = 0 Then Line2(l).BorderColor = &H60& Else Line2(l).BorderColor = &HFF&
            Case 2: If s = 0 Then Line3(l).BorderColor = &H60& Else Line3(l).BorderColor = &HFF&
        End Select
    Next l

Next k
End Sub

Private Sub Form_Load()
Call Digit(HScroll1.Value)
End Sub

Private Sub HScroll1_Change()
Call Digit(HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
Call Digit(HScroll1.Value)
End Sub
