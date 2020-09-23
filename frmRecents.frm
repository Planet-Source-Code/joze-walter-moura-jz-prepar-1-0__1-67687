VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecents
   BackColor       =   &H00FFFFFF&
   Caption         =   " Part 1 :  RECENT FILES"
   ClientHeight    =   6015
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen
      Caption         =   "Open a File"
      Height          =   525
      Left            =   5625
      TabIndex        =   1
      ToolTipText     =   "Sample Opening File"
      Top             =   45
      Width           =   1335
   End
   Begin VB.CommandButton cmdMore
      Caption         =   "More ... Part 2"
      Height          =   510
      Left            =   5520
      TabIndex        =   0
      ToolTipText     =   "Open Parameters & Preferences form"
      Top             =   2505
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog CDg
      Left            =   5100
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4
      Alignment       =   2  'Center
      Caption         =   "SUGGESTED TEST:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   6390
   End
   Begin VB.Label Label3
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRecents.frx":0000
      BeginProperty Font
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1935
      Left            =   3630
      TabIndex        =   5
      Top             =   3690
      Width           =   3165
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRecents.frx":00C8
      BeginProperty Font
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1935
      Left            =   330
      TabIndex        =   4
      Top             =   3690
      Width           =   3165
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1
      BorderColor     =   &H0000C000&
      Height          =   2115
      Left            =   135
      Shape           =   4  'Rounded Rectangle
      Top             =   3615
      Width           =   6810
   End
   Begin VB.Label lblFile
      Alignment       =   2  'Center
      Caption         =   "Working File"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   6960
   End
   Begin VB.Label Label2
      BackStyle       =   0  'Transparent
      Caption         =   "Working Selected File"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2790
      TabIndex        =   2
      Top             =   630
      Width           =   1560
   End
   Begin VB.Menu mnufil
      Caption         =   "&File"
      Begin VB.Menu mnusep_a
         Caption         =   "-"
      End
      Begin VB.Menu mnufilrec
         Caption         =   "&Recent Files"
         Begin VB.Menu mnuRecentFiles
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnusep_z
         Caption         =   "-"
      End
      Begin VB.Menu mnufilexi
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRecents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'.----------------------------------------------------------------------
'| Module    : frmRecents - Template to include and use RECENT FILES
'| About     : 19/01/2007 12:17 -o- Author: JOZE Walter de Moura
'| Credits   : Thank's a lot to several authors at PSC for about 20% of
'|             all code I've used in this work.
'`======================================================================
'| Purpose   : Remember last used FileNames opened by your App.
'| Advantage : No Windows Register - will be a INI file where you want.
'|             Auto exclude obsolete files, Shorten view enchance, etc.
'`----------------------------------------------------------------------
'  How to:
'    a) use your-form Menu Editor, create any as "mnuRecentFiles", Index=0
'    b) Include modPrePar.Bas module in your Project
'    c) See at few highlighted code lines to be your form inserted
'       in some essentials as Form_Load, Form_Unload, getting a CDlg path,
'       menu clicking, etc.
'
'  Enjoy, Joze from Rio de Janeiro, Brazil.
'
Option Explicit

Private egLoadSpec As String ' e.g A file as open
                                       ' recent files appearing in Menu

' caller for 2nd sample
Private Sub cmdMore_Click()
   frmPrePar.Show
   Unload Me
End Sub


Private Sub Form_Load()
'.- - - - - - - - - - - - - - - - - - - - - - - - - - - -
    'INI File Name for Parameters ... would be any file.INI, any dir
    JzIniFile = App.Path & "\Sample.Ini" 'e.g.
    'for Recent List working
    ReDim RecentsArray(1) ' array prepare
    JzGetRecentsFromINI mnuRecentFiles ' load previous recent list
'`- - - - - - - - - -  - - - - - - - - - - - - - - - - - -
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = False 'Inihibit
'.- - - - - - - - - - - - - - - - - - - - - - - - -
        JzToINIFile ' Saves actual Recent list to INIFile
'`- - - - - - - - - - - - - - - - - - - - - - - - -
        Unload Me
    End If
End Sub

Private Sub mnufilexi_Click()
'.- - - - - - - - - - - - - - - - - - - - - - - - -
    JzToINIFile ' Saves actual Recent list to INIFile
'`- - - - - - - - - - - - - - - - - - - - - - - - -
    Unload Me
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
'.- - - - - - - - - - - - - - - - - - - - - - - - - - -
   If Len(Trim(mnuRecentFiles(Index).Caption)) > 0 Then
      egLoadSpec = JzFromRecents(Index)
      'E.G. NOW YOU HAVE
      '     THE CHOOSED FILENAME TO WORK ETC
      'Our Sample process is:
      lblFile.Caption = egLoadSpec ' Show
   End If
'`- - - - - - - - - - - - - - - - - - - - - - - - - - -
End Sub

'E.g. Getting a FileSpec from a Common Dialog
Private Sub CmdOpen_Click()
Dim k As Long
    MousePointer = vbDefault
    With CDg
        .Flags = cdlOFNFileMustExist
        .CancelError = True 'to Cancel effect
        .DialogTitle = " File Open Sample:"
'.- - - - - - - - - - - - - - - - - - - - - - - - - - -
        .InitDir = JzFromRecents(0) ' 1st = +recent
'`- - - - - - - - - - - - - - - - - - - - - - - - - - -
        .Filter = "All File Types |*.*|e.g. Some Graphic Files|*.jpg;*.bmp;*.gif"
    End With
    On Error Resume Next
    CDg.ShowOpen
    If Not Err Then
       egLoadSpec = CDg.FileName
'.- - - - - - - - - - - - - - - - - - - - - - - - - - -
       If JzToRecents(egLoadSpec, mnuRecentFiles) Then ' put/get
             'E.G. NOW YOU HAVE
             '     THE CHOOSED FILENAME TO WORK ETC
             'Our Sample process is:
             lblFile.Caption = egLoadSpec ' Show it
       End If
'`- - - - - - - - - - - - - - - - - - - - - - - - - - -
    End If
    On Error GoTo 0
End Sub

'-oOo-oOo-oOo-

