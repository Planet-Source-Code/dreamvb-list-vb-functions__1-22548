VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   330
      Left            =   5730
      TabIndex        =   4
      Top             =   90
      Width           =   810
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   330
      Top             =   885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   330
      Left            =   5205
      TabIndex        =   1
      Top             =   90
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   5025
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   6420
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   5565
      Width           =   6450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data As String
Function FindEndFunc(S As String) As Integer
    X = InStr(S, "End Function")
    FindEndFunc = X
    
End Function
Function FindFunc(S As String) As Integer
    X = InStr(S, "Function")
    FindFunc = X
    
End Function
Function FindBrace(S As String) As Integer
    X = InStr(S, "(")
    If X = 0 Then Exit Function
    FindBrace = X
    
End Function
Function ListFunctionNames(StrData As String)
Dim I As Integer
Dim V As Variant
Dim Strline As String
Dim X, Y As Integer
Dim FuncCount As Integer
On Error Resume Next
    
    V = Split(StrData, vbCrLf)
    '////////////////////////////////////////////
    For I = LBound(V) To UBound(V)
        Strline = Trim(V(I))
        X = FindFunc(Strline)
        Y = FindBrace(Strline)
        If X = 0 Then
            ElseIf Y = 0 Then
        Else
            FuncCount = FuncCount + 1
            Combo1.AddItem Mid(Strline, X + 9, Y - X - 9)
          End If
        Next
    '////////////////////////////////////////////
    Label1.Caption = FuncCount & " Functions found in file."
    
End Function

Private Sub Combo1_Click()
Dim FindFunc As String
    Text1 = ""
    FindFunc = Trim(Combo1.Text)
        X = InStr(Data, FindFunc)
        Y = FindEndFunc(Data)
        A = Mid(Data, X, Y + 4)
        X = 0
        Y = 0
        X = InStr(A, "End Function")
        Text1 = "Function " & Mid(A, 1, X + 11)
        
End Sub

Private Sub Command1_Click()
Dim Fnum As Long
    Combo1.Clear
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.FileName) = 0 Then Exit Sub
    '///////////////////////////////////////////////
    Open CommonDialog1.FileName For Binary As #1
        Data = Space(LOF(1))
            Get #1, , Data
        Close #1
    ListFunctionNames Data
    '///////////////////////////////////////////////
End Sub

Private Sub Command2_Click()
    Unload Form1: End
End Sub
