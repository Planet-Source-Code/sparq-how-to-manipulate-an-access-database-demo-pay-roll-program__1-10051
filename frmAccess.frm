VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sparq's Access Database  -  "
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   1680
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3180
      TabIndex        =   12
      Top             =   1680
      Width           =   1155
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   1680
      Width           =   1155
   End
   Begin VB.TextBox txtSSNum 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   870
      Width           =   2115
   End
   Begin VB.TextBox txtPayRate 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1170
      Width           =   1995
   End
   Begin VB.TextBox txtEmpNum 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   570
      Width           =   2115
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   270
      Width           =   2115
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3300
      Top             =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   1740
      X2              =   6840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1740
      X2              =   6840
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   1680
      X2              =   6780
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1680
      X2              =   6780
      Y1              =   2715
      Y2              =   2715
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pot Hedds Inc.                                                            "
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   1860
      TabIndex        =   15
      Top             =   2760
      Width           =   3795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   5
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SS Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   4
      Top             =   900
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   3
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   2
      Top             =   300
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employees"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Db As Database
Public EmployeeTable As Recordset

Public EditText As Boolean
Public AddUser As Boolean
Public LastSpot As Integer

Private Function UpdateUser()
  Dim Spot As Integer
    Spot = InStr(1, txtName, " ")
    With EmployeeTable
        If AddUser = False Then
            .MoveFirst
            Do Until List1.Text = !FName & " " & !Lname
                .MoveNext
            Loop

            .Edit
        Else
            .AddNew
        End If
            !FName = Left(txtName, Spot - 1)
            !Lname = Right(txtName, Len(txtName) - Spot)
            !EmpNum = txtEmpNum
            !SSNum = txtSSNum
            !PayRate = txtPayRate
            .Update
    End With
    LoadNames
End Function


Private Sub cmdAddNew_Click()
    LastSpot = List1.ListIndex
    If Left(cmdAddNew.Caption, 1) = "A" Then
        cmdAddNew.Caption = "Update"
        EditText = True
        
        txtName = ""
        txtEmpNum = ""
        txtSSNum = ""
        txtPayRate = ""
        
    Else
        cmdAddNew.Caption = "Add New"
        EditText = False
        AddUser = True
        UpdateUser
    End If
    
    txtName.Locked = Not (EditText)
    txtEmpNum.Locked = Not (EditText)
    txtSSNum.Locked = Not (EditText)
    txtPayRate.Locked = Not (EditText)
    List1.Enabled = Not (EditText)
    cmdEdit.Enabled = Not (EditText)
    cmdDelete.Enabled = Not (EditText)
    cmdCancel.Enabled = EditText
    cmdAddNew.Enabled = EditText
End Sub

Private Sub cmdCancel_Click()
    EditText = False
    List1.Enabled = Not (EditText)
    List1.ListIndex = LastSpot
    List1_Click
    
    txtName.Locked = Not (EditText)
    txtEmpNum.Locked = Not (EditText)
    txtSSNum.Locked = Not (EditText)
    txtPayRate.Locked = Not (EditText)
    cmdEdit.Caption = "Edit"
    cmdAddNew.Caption = "Add New"
    cmdCancel.Enabled = EditText
    cmdDelete.Enabled = Not (EditText)
End Sub


Private Sub cmdDelete_Click()
    Dim Answer As Boolean
    
    Answer = MsgBox("Delete " & List1.Text & "?", vbQuestion + vbYesNo, "Confirm")
    If Answer = vbYes Then DeleteUser
End Sub


Private Function DeleteUser()
    With EmployeeTable
           .MoveFirst
           Do Until txtName = !FName & " " & !Lname
               .MoveNext
           Loop
           .Delete
    End With
End Function


Private Sub cmdEdit_Click()
    LastSpot = List1.ListIndex
    If Left(cmdEdit.Caption, 1) = "E" Then
        cmdEdit.Caption = "Update"
        EditText = True
    Else
        cmdEdit.Caption = "Edit"
        EditText = False
        AddUser = False
        UpdateUser
    End If
    
    txtName.Locked = Not (EditText)
    txtEmpNum.Locked = Not (EditText)
    txtSSNum.Locked = Not (EditText)
    txtPayRate.Locked = Not (EditText)
    List1.Enabled = Not (EditText)
    cmdDelete.Enabled = Not (EditText)
    cmdCancel.Enabled = EditText
End Sub


Private Sub Form_Load()
    
   'Initialize BOTH the Database and the Table
    Set Db = OpenDatabase(App.Path & "\AccessDB.mdb")
    Set EmployeeTable = Db.OpenRecordset("Employees", dbOpenDynaset)
    
    LoadNames  'Fill ListBox
    
End Sub


Private Function LoadNames()   'Fill Listbox
    List1.Clear
    With EmployeeTable
        .MoveFirst
        Do While Not .EOF
            List1.AddItem !FName & " " & !Lname
            .MoveNext
        Loop
    End With
End Function


Private Function GetUserInfo(FullName As String)
    With EmployeeTable
        .MoveFirst
        Do Until txtName = !FName & " " & !Lname
            .MoveNext
        Loop
        txtEmpNum = !EmpNum
        txtSSNum = !SSNum
        txtPayRate = !PayRate
    End With
End Function


Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Vote for me if you'd like." & vbCrLf & _
           "Keep Puffing"
End Sub

Private Sub List1_Click()
    txtName = List1.List(List1.ListIndex)
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    GetUserInfo (List1.Text)
End Sub


Private Sub Timer1_Timer()
  Dim String1 As String
  Dim String2 As String
    String1 = Left(Caption, 1)
    String2 = Right(Caption, Len(Caption) - 1)
    Caption = String2 & String1

    String1 = Right(Label7, 1)
    String2 = Left(Label7, Len(Label7) - 1)
    Label7 = String1 & String2
End Sub
