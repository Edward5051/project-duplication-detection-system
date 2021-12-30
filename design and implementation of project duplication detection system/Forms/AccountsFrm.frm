VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form AccountsFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMISSION NUMBER: 1710207040; NAME: WAZIRI JOB"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9525
   ClipControls    =   0   'False
   Icon            =   "AccountsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00D8E9EC&
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   9540
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   9015
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView lstUser 
      CausesValidation=   0   'False
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   128
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1023
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   3528
      EndProperty
   End
   Begin trapduplicate.jcbutton cmdAdd 
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   7080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "&Add"
      UseMaskCOlor    =   -1  'True
   End
   Begin trapduplicate.jcbutton cmdCancel 
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   7080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "&Cancel"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervised by: Dr Nurudeen Ibrahim "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   5715
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Duplication Detection "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4710
   End
End
Attribute VB_Name = "AccountsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsData As New clsUsers

Private Sub cmdClose_Click()
Unload Me
End Sub



Private Sub cmdAdd_Click()
    If cmdAdd.Caption = "&Add" Then
        cmdAdd.Caption = "&Save"
        Lockfalse
        Clear
    Else
        If ValidateEntry = False Then
        cmdAdd.Caption = "&Add"
        txtUser.SetFocus
        Lockfalse
        clsData.AddUser txtUser
        Clear
        Unload Me
        AccountsFrm.Show
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub







Private Sub Form_Activate()
Set clsData = New clsUsers
clsData.DisplayUsers lstUser
End Sub

Public Sub Lockfalse()
txtUser.Locked = False

End Sub

Public Sub Clear()
txtUser = ""


End Sub
Public Function ValidateEntry() As Boolean
ValidateEntry = True
If Trim(txtUser) = "" Then
    MsgBox "Please enter username", vbInformation, ""
    Exit Function
End If
ValidateEntry = False
End Function

