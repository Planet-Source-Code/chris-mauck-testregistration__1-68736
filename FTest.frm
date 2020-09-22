VERSION 5.00
Begin VB.Form FTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Registration"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear All"
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame0 
      Caption         =   "GENERATE DYNAMIC REGISTRATION KEY"
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton cmdCreate 
         Caption         =   "CREATE"
         Height          =   375
         Left            =   4320
         TabIndex        =   24
         Top             =   2610
         Width           =   860
      End
      Begin VB.TextBox txtDateSpecific 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtAppSpecific 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Text            =   "PP2007CM"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Clicking on Create will display the number that could be displayed when your registration form is loaded."
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   $"FTest.frx":0000
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   5175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hidden Variable Information"
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   320
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "{"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   735
         Left            =   1200
         TabIndex        =   25
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblOwner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label lblOwner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   960
         TabIndex        =   22
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label lblOwner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   21
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label lblOwner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   2640
         TabIndex        =   20
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label lblOwner 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   3480
         TabIndex        =   19
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Dynamically Generated Date:"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   675
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Application Specific Characters"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GENERATE ACTIVATION KEY FROM REGISTRATION KEY"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   5415
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "GEN"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   330
         Width           =   860
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "PLEASE ENTER PRODUCT ACTIVATION KEY"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   5415
      Begin VB.CommandButton cmdTest 
         Caption         =   "TEST"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   330
         Width           =   860
      End
      Begin VB.TextBox txtAlpha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   5
         Top             =   360
         Width           =   760
      End
      Begin VB.TextBox txtAlpha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   4
         Top             =   360
         Width           =   760
      End
      Begin VB.TextBox txtAlpha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   3
         Top             =   360
         Width           =   760
      End
      Begin VB.TextBox txtAlpha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   960
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   760
      End
      Begin VB.TextBox txtAlpha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   760
      End
   End
End
Attribute VB_Name = "FTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FTest
' AUTHOR:       Phil Fresle
' CREATED:      06-Sep-2000
' COPYRIGHT:    Copyright 2000 Frez Systems Limited.
'
' PRESENT CONCEPT:  Chris Mauck
'
' Original program created by Phil Fresle and submitted to freevbcode.com
'
' Modified 6-03-2007 by Chris Mauck to be used as registration generator.
'
' Nothing special has been added to the program by me. I just modified the
' display, took out the unecessary routines (for my purpose) and put it an
' easy to understand 'tutorial' style application.
'
' This is the first portion of the program and the number that is generated
' by the program can be used to call in to support to have a KEY generated
' and returned to the user for access OR in my case I am developing a Perl
' script that can generate the same activation KEY through a website form.
'
'*******************************************************************************
Option Explicit

Private Sub cmdClear_Click()
    Dim i As Integer
    
    For i = 0 To 4
        lblOwner(i).Caption = ""
        lblTest(i).Caption = ""
        txtAlpha(i).Text = ""
    Next i
    
    txtDateSpecific.Text = ""
    
    Call DynDate
End Sub

'*******************************************************************************
' Test key generation for generic alpha numeric keys
' - This uses the Product Specific text (set by the developer and hidden in actual use)
' - As well as the date, or whatever field you would choose to add in.
' The result is the key like you would see on the Microsoft registrations before
' you would call in to get yet another code.
' For my use I will be creating the Perl script to accomodate the same function
' as the "cmdGenerate_Click" sub. cmdGenerate_Click's results will be hidden and
' users will have to register the result of cmdCreate_Click at my website to get
' their final KEY.
' For our purpose here, cmdGenerate_Click just displays what we should type into
' the KEY code text boxes.
'*******************************************************************************
Private Sub cmdCreate_Click()
    Dim oReg As CGenericRegistration
    Dim sKey As String
    
    Set oReg = New CGenericRegistration
    
    sKey = oReg.GenerateKey(txtAppSpecific.Text & txtDateSpecific.Text)
    
    lblOwner(0).Caption = Left(sKey, 5)
    lblOwner(1).Caption = Mid(sKey, 6, 5)
    lblOwner(2).Caption = Mid(sKey, 11, 5)
    lblOwner(3).Caption = Mid(sKey, 16, 5)
    lblOwner(4).Caption = Mid(sKey, 21, 5)
        
    Set oReg = Nothing
End Sub

'*******************************************************************************
' Activation KEY generation - sample display for test purposes.
' For our purpose here, cmdGenerate_Click just displays what we should type into
' the KEY code text boxes.
'*******************************************************************************
Private Sub cmdGenerate_Click()
    Dim oReg As CGenericRegistration
    Dim sKey As String
    Dim cKey As String
    
    Set oReg = New CGenericRegistration
    
    cKey = lblOwner(0).Caption & _
           lblOwner(1).Caption & _
           lblOwner(2).Caption & _
           lblOwner(3).Caption & _
           lblOwner(4).Caption
    
    sKey = oReg.GenerateKey(cKey)
    
    lblTest(0).Caption = Left(sKey, 5)
    lblTest(1).Caption = Mid(sKey, 6, 5)
    lblTest(2).Caption = Mid(sKey, 11, 5)
    lblTest(3).Caption = Mid(sKey, 16, 5)
    lblTest(4).Caption = Mid(sKey, 21, 5)
        
    Set oReg = Nothing
End Sub

'*******************************************************************************
' Test key validation
' - This is where your program would check for validity and grant access.
'*******************************************************************************
Private Sub cmdTest_Click()
    Dim sKey As String
    Dim cKey As String
    Dim oReg As CGenericRegistration
    
    sKey = txtAlpha(0).Text & _
           txtAlpha(1).Text & _
           txtAlpha(2).Text & _
           txtAlpha(3).Text & _
           txtAlpha(4).Text
    cKey = lblOwner(0).Caption & _
           lblOwner(1).Caption & _
           lblOwner(2).Caption & _
           lblOwner(3).Caption & _
           lblOwner(4).Caption
    
    Set oReg = New CGenericRegistration
    
    If oReg.IsKeyOK(sKey, cKey) Then
        MsgBox "Product activated successfully!!" & vbCrLf & vbCrLf & _
               "Thank you for choosing PRODUCT NAME." & vbCrLf & _
               "We appreciate your support!", vbOKOnly, "Product Registration"
    Else
        MsgBox "Product key is BAD!!" & vbCrLf & vbCrLf & _
               "Please ensure that you have entered the" & vbCrLf & _
               "key exactly as it appeared within your" & vbCrLf & _
               "registration confirmation." & vbCrLf & vbCrLf & _
               "Please try again.", vbCritical, "Product Registration"
    End If
    
    Set oReg = Nothing
End Sub

Private Sub Form_Load()
    Call DynDate
End Sub

Private Sub txtAlpha_Change(Index As Integer)
    On Error GoTo cmdTestFocus
    With txtAlpha(Index)
        If Len(.Text) = .MaxLength Then
            .Text = UCase(.Text)
            txtAlpha(Index + 1).SetFocus
        End If
    End With
    
cmdTestFocus:
    If Err Then cmdTest.SetFocus
End Sub

Private Sub txtAlpha_GotFocus(Index As Integer)
    txtAlpha(Index).SelStart = 0
    txtAlpha(Index).SelLength = Len(txtAlpha(Index).Text)
End Sub

'*******************************************************************************
' Dynamic string.
' - Take the current system date and time to be added to the end of the pre-set
'   'AppSpecific' text before temporary key is generated.
'*******************************************************************************
Public Sub DynDate()
    Dim MyDate, MyTime, MyMonth, MyDay, MyYear, MyHour, MyMin

    MyDate = Date
    MyTime = Time
    MyMonth = Month(MyDate)
        If Len(MyMonth) = 1 Then MyMonth = "0" & MyMonth
    MyDay = Day(MyDate)
        If Len(MyDay) = 1 Then MyDay = "0" & MyDay
    MyYear = Year(MyDate)
    MyHour = Hour(MyTime)
        If Len(MyHour) = 1 Then MyHour = "0" & MyHour
    MyMin = Minute(MyTime)
        If Len(MyMin) = 1 Then MyMin = "0" & MyMin
    
    txtDateSpecific.Text = MyHour & MyMin & MyMonth & MyDay & MyYear
End Sub
