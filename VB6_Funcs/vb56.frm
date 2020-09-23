VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " Compile first"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Case insensitive"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1770
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   1950
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   180
      Width           =   5085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Time test"
      Height          =   525
      Left            =   330
      TabIndex        =   0
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim s As String
    Dim sA() As String
    Dim vbMethod As VbCompareMethod
    
    Dim i As Long
    Dim d As Double

    Text1 = "For i = 1 To 100000" & vbCrLf
    Text1.SelStart = Len(Text1)
    Text1.Refresh

    s = "a b c def ghijk  lmn o p q r s t u v wxy z"
    vbMethod = Check1.Value

        ''''''''''''''''
        ' VB6 substitue Split function
        d = ProfileStart
        For i = 1 To 100000
            sA = VB6_Funcs.Split(s, , 13, vbMethod, 1)
        Next
        d = ProfileStop(d, epSeconds)
        ''''''''''''''''

    Text1.SelText = vbCrLf
    Text1.SelText = "sA = VB6_Funcs.Split(s, , 13, , 1)" & vbCrLf
    Text1.SelText = FormatElapsed(d, 4) & " seconds" & vbCrLf
    Text1.SelText = "sA(LBound(" & LBound(sA) & ")) = " & sA(LBound(sA)) & vbCrLf
    Text1.SelText = "sA(UBound(" & UBound(sA) & ")) = " & sA(UBound(sA)) & vbCrLf

        ''''''''''''''''
        ' VB5 Split sub-routine
        d = ProfileStart
        For i = 1 To 100000
            VB6_Funcs.SplitVb5 s, sA, , 13, vbMethod
        Next
        d = ProfileStop(d, epSeconds)
        ''''''''''''''''
    
    Text1.SelText = vbCrLf
    Text1.SelText = "VB6_Funcs.SplitVb5 s, sA, , 13" & vbCrLf
    Text1.SelText = FormatElapsed(d, 4) & " seconds" & vbCrLf
    Text1.SelText = "sA(LBound(" & LBound(sA) & ")) = " & sA(LBound(sA)) & vbCrLf
    Text1.SelText = "sA(UBound(" & UBound(sA) & ")) = " & sA(UBound(sA)) & vbCrLf

        ''''''''''''''''
        ' VB6 Split function
        d = ProfileStart
        For i = 1 To 100000
            sA = VBA.Strings.Split(s, , 13, vbMethod)
        Next
        d = ProfileStop(d, epSeconds)
        ''''''''''''''''

    Text1.SelText = vbCrLf
    Text1.SelText = "sA = VBA.Strings.Split(s, , 13)" & vbCrLf
    Text1.SelText = FormatElapsed(d, 4) & " seconds" & vbCrLf
    Text1.SelText = "sA(LBound(" & LBound(sA) & ")) = " & sA(LBound(sA)) & vbCrLf
    Text1.SelText = "sA(UBound(" & UBound(sA) & ")) = " & sA(UBound(sA)) & vbCrLf

        ''''''''''''''''
        ' VB5 Join function
        d = ProfileStart
        For i = 1 To 100000
            s = VB6_Funcs.Join(sA, "")
        Next
        d = ProfileStop(d, epSeconds)
        ''''''''''''''''
    
    Text1.SelText = vbCrLf
    Text1.SelText = "VB6_Funcs.Join    " & FormatElapsed(d, 4) & " seconds" & vbCrLf
    Text1.SelText = s & vbCrLf
    
        ''''''''''''''''
        ' VB6 Join function
        d = ProfileStart
        For i = 1 To 100000
            s = VBA.Strings.Join(sA, "")
        Next
        d = ProfileStop(d, epSeconds)
        ''''''''''''''''
    
    Text1.SelText = vbCrLf
    Text1.SelText = "VBA.Strings.Join  " & FormatElapsed(d, 4) & " seconds" & vbCrLf
    Text1.SelText = s & vbCrLf

        ''''''''''''''''
        ' VB6 Join$ function
        d = ProfileStart
        For i = 1 To 100000
            s = VBA.Strings.Join$(sA, "")
        Next
        d = ProfileStop(d, epSeconds)
        ''''''''''''''''
    
    Text1.SelText = vbCrLf
    Text1.SelText = "VBA.Strings.Join$ " & FormatElapsed(d, 4) & " seconds" & vbCrLf
    Text1.SelText = s & vbCrLf

End Sub
