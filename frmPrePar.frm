VERSION 5.00
Begin VB.Form frmPrePar 
   BackColor       =   &H00808080&
   Caption         =   "Part 2: PARAMETERS & PREFERENCES"
   ClientHeight    =   6015
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   150
      Picture         =   "frmPrePar.frx":0000
      ScaleHeight     =   2865
      ScaleWidth      =   2280
      TabIndex        =   1
      Top             =   150
      Width           =   2310
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Return ... Part 1"
      Height          =   510
      Left            =   5400
      TabIndex        =   0
      ToolTipText     =   "Open Recent Files form"
      Top             =   195
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "SUGGESTED TEST:"
      Height          =   195
      Left            =   330
      TabIndex        =   4
      Top             =   3375
      Width           =   6390
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPrePar.frx":7A19
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1920
      Left            =   3585
      TabIndex        =   3
      Top             =   3735
      Width           =   3165
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPrePar.frx":7AD1
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1920
      Left            =   300
      TabIndex        =   2
      Top             =   3735
      Width           =   3165
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFC0&
      Height          =   2115
      Left            =   105
      Shape           =   4  'Rounded Rectangle
      Top             =   3630
      Width           =   6810
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuUsr 
         Caption         =   ""
      End
      Begin VB.Menu mnuExi 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEna 
      Caption         =   "&Enabled Operations"
      HelpContextID   =   6
      Begin VB.Menu mnuEnaDel 
         Caption         =   "&Delete e.g.Anything"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEnaTrf 
         Caption         =   "&Transfer e.g.Value"
      End
      Begin VB.Menu mnuEnaEmp 
         Caption         =   "&Empt e.g.Tables"
      End
   End
   Begin VB.Menu mnuChk 
      Caption         =   "&Checked Preferences"
      HelpContextID   =   7
      Begin VB.Menu mnuChkDnt 
         Caption         =   "&Don't e.g. Clear after Write"
      End
      Begin VB.Menu mnuChkWhi 
         Caption         =   "&Change Color Schemme"
      End
      Begin VB.Menu mnuChkHid 
         Caption         =   "Show/Hide (beauty) &Picture"
      End
      Begin VB.Menu mnuChkLod 
         Caption         =   "&Load e.g. At Start"
      End
   End
   Begin VB.Menu mnuVar 
      Caption         =   "&Variables"
      Begin VB.Menu mnuVarTim 
         Caption         =   "Default Initial &Time=9:00"
      End
      Begin VB.Menu mnuVarSig 
         Caption         =   "&Who Signs Doccument=ME"
      End
   End
End
Attribute VB_Name = "frmPrePar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'.----------------------------------------------------------------------
'| Module    : frmPrePar -Template to include and use Parameters on Menu
'| About     : 19/01/2007 12:17 -o- Author: JOZE Walter de Moura
'| Credits   : Thank's a lot to several authors at PSC for about 20% of
'|             all code I've used in this work.
'`======================================================================
'| Purpose   : App Settings at form(s) Menu saved/loaded with a INI File
'| Advantage : No Windows Register, No System Vars, minimum code effort
'`----------------------------------------------------------------------
'  How to:
'    a) use your-form Menu Editor, create about settings types you need:
'           - Checked/Unchecked user preferences;
'           - Enabled/Disabled user permissions;
'           - User/Install pre-fixed Variables;
'           - Customized Caption in Menu;
'           - Others your idea - visible menu itens, etc.
'
'    b) Include modPrePar.Bas module in your Project

'    c) See at few highlighted code lines to be your form inserted
'       in some essentials as Form_Load'.
'
'    d) Examine code treatment as sample and aply similar to your App.
'
'  Enjoy, Joze from Rio de Janeiro, Brazil.
'
'Pre Parameters Initialized Variables
Private egDefSigner As String 'e.g. A person to sign Doccuments (ME)
Private egDefTime As String 'e.g. At what time to Start (9:00)

'------------------ Internals for Sample -----------------------
'Exiting by Menu
Private Sub mnufilexi_Click()
  Unload Me
End Sub

'Returning Previous Form
Private Sub cmdMore_Click()
  frmRecents.Show
  Unload Me
End Sub

'This is to simulate a variable form color schemm
'(poor, only to illustrate)
Private Sub egChangeColors()
   If mnuChkWhi.Checked Then 'no white
      Me.BackColor = &H808080
      Shape1.BorderColor = &HFFFFC0
      Label1.ForeColor = &H80FFFF
      Label2.ForeColor = &H80FFFF
   Else
      Me.BackColor = &HFFFFFF 'white
      Shape1.BorderColor = &HC000&
      Label1.ForeColor = &H404040
      Label2.ForeColor = &H404040
   End If
End Sub

'================== USE AS THIS - CUT, PAST, MODIFY RESPECTIVE CODE ==============

Private Sub Form_Load()
  Me.Caption = " Part 2: PARAMETERS ON MENU"
'.------------------------------------------------------
'INIFile ID
   'INI File Name for Parameters ... would be any file.INI, any dir
   JzIniFile = App.Path & "\Sample.Ini"

   'Getting Caption for a Menu Item  (direct code)
   JzSection = "Customized Captions"
   JzEntry = "mnuUsr"
   JzValue = "Customized or Translated Caption" ' default value
   JzRemarks = JzValue ' remarks = original Menu item Caption
   JzGetINI
   mnuUsr.Caption = JzValue
'to test above, do NotePad Sample.Ini and
'change "mnuUsr" entry to a new phrase
'so execute me again and see it

   'Getting Check Status from INI file
   JzParCheck mnuChkDnt
   JzParCheck mnuChkWhi
   JzParCheck mnuChkLod
   JzParCheck mnuChkHid
'doing Forms adjusts due preferences
   Picture1.Visible = Not mnuChkHid.Checked
   'restabelish previous color schemm
   egChangeColors

   'Getting Enabled Status from INI file
   JzParEnable mnuEnaDel
   JzParEnable mnuEnaTrf
   JzParEnable mnuEnaEmp

   'Pre Initialized Variables
   egDefTime = JzParVariable(mnuVarTim, "9:00")
   egDefSigner = JzParVariable(mnuVarSig, "ME")
'`---------------------------------------------------
End Sub

'.---------------------------------------------------
' CHECKED PARAMETERS WILL BE JUST UPDATED
Private Sub mnuChkHid_Click()
  If Not mnuChkHid.Checked Then
     mnuChkHid.Checked = True
  Else
     mnuChkHid.Checked = False
  End If
  JzCheckToINI mnuChkHid
  Picture1.Visible = Not mnuChkHid.Checked
End Sub

Private Sub mnuChkDnt_Click()
  If Not mnuChkDnt.Checked Then
     mnuChkDnt.Checked = True
  Else
     mnuChkDnt.Checked = False
  End If
  JzCheckToINI mnuChkDnt
  '
  'Do Anything about Parameters Checked status
  '
End Sub

Private Sub mnuChkWhi_Click()
  If Not mnuChkWhi.Checked Then
     mnuChkWhi.Checked = True
  Else
     mnuChkWhi.Checked = False
  End If
  egChangeColors ' e.g. a schemme option
  JzCheckToINI mnuChkWhi
  Me.Refresh
End Sub

Private Sub mnuChkLod_Click()
  If Not mnuChkLod.Checked Then
     mnuChkLod.Checked = True
  Else
     mnuChkLod.Checked = False
  End If
  JzCheckToINI mnuChkLod
  '
  'Do Anything about Parameters Checked status
  '
End Sub
'.................. End Code for Checked  ..........................

'.------------- IRRELEVANT CODE -------------------------------
'These Pseudo Procedures is only for code examples aplying
'Checked Menu Parameters

Private Sub egAppAction()
End Sub

Private Sub egDoALotOfNothing()
Dim eg_dataf As String
     If eg_dataf = "x" And mnuChkLod.Checked = True Then
        egAppAction
     End If
End Sub
Private Sub DoOperation()
     If mnuEnaEmp.Enabled Then
        'do it
     End If
End Sub
'................. End Irrelevant Code ............................

'.---------------------------------------------------
' VARIABLES PARAMETERS UPDATED VIA INPUTBOX
'user want modify variables
Private Sub mnuVarTim_Click()
  egDefTime = JzNewVariable(mnuVarTim)
End Sub

Private Sub mnuVarSig_Click()
  egDefSigner = JzNewVariable(mnuVarSig, _
      "Who will sign Company Doccuments?")
End Sub
'.................. End code for Variables .................................

'.---------------------------------------------------
' ENABLED PARAMETERS E.G., NO MODIFICATIONS HERE
' General uses is about light installation permissions
' once configurated - may be via NotePad Sample.INI
' or if you want provide a Administrator Form to do it
'
' Users will not click on it if disabled:

Private Sub mnuEnaDel_Click()
   MsgBox "e.g. a Delete Operation" 'Admin
End Sub

Private Sub mnuEnaTrf_Click()
   MsgBox "e.g. a Transfer Operation" 'Admin
End Sub

Private Sub mnuEnaEmp_Click()
   MsgBox "e.g. a Cleanner Operation" 'Advanced User
End Sub
'................. End Code for Enabled ...........................

'-oOo-oOo-oOo-
