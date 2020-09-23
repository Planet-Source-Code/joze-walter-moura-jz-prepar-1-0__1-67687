Attribute VB_Name = "modPrePar"
'.---------------------------------------------------------------------------------
'| This module is not all from me - 20% was from several PSC freecode whose authors
'| come in with his credits. The idea is to concentre in this the bulk programming
'| hard code to use Parameters, User Preferences and/or Recent Files in a App forms
'| Menus with Set/Reset/Updating features by including a few code lines.
'| See PrePar.Txt and examine frmRecents and frmPrePar for easy template. [Joze]
'`---------------------------------------------------------------------------------
Option Explicit
'.---------------------------------------------------------------------------------
'| Code Features by Subject
'| a) General INI File use: pre-defined Public parameters; remarks in same entry
'|                          line; read, write and especial get operation that try
'|                          read - if fail then create all necessary.
'| b) Recent Files: auto updating inclusive excluding if file no more exists; long
'|                  filespecs are optimized for user menu view, w/variable string
'|                  size(=30); Update operation reforms queue in INI File; variable
'|                  maximum number of showed files.
'| c) Preferences and Parameters: Each Menu item will be one of this, if designed.
'|                                Checked boolean logic, Enabled/Visible install
'|                                user Parameters, Preseting Variable - these are
'|                                code treated; You may do more you want - template
'|                                illustrates as sample a Customizable Menu item
'|                                Caption using hard code;
'|                                Check/Uncheck may be user on line updated;
'|                                Variables, when clicked on, will be kbd accepted.
'`---------------------------------------------------------------------------------
' This Module was coded mode you can cut something you don't need your app, e.g.,
' Recent Files feature, etc.
' If you find bugs or do enchancements, please, let me know, thank you! [Joze]
'                                                              jozew@globo.com
' 16/01/2007

' Independent I-O INI parameters elements
Public JzIniFile As String 'Complete path with INI file specs
Public JzSection As String 'Working INI Section name (without []s)
Public JzEntry As String ' Working Key Name leading "=" sign
Public JzValuINI As String ' String all after "=" sign (JzValue + " ; " + JzRemarks)
Public JzValue As String ' Value string 1st part of JzValuINI (until last ";")
                         ' contains real value to i-o operations.
Public JzRemarks As String ' Remarks string 2nd part of JzValuINI (after last ";")
' internal
Private PrevINIFile As String ' Break Control for file tests

'Recent File Works - modify its values if required
Public Const MaxRecentFiles As Integer = 4 ' Preference for Recent List feature
Public Const ShortLen As Integer = 30 ' Maximum shorten strings size in Menu Recents
Public RecentsArray() As String ' Work area for Recent List feature

'INI General
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long

' lpDefault must be setting with default value to be returned if operation fails
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

'.-----------------------------------------------------------------------------
'| INI Operations : all of this requires pre-inicialization of
'|     JzIniFile, JzSection and JzEntry.
'|     Others (JzValue, JzValuINI and JzRemarks) are process dependents.
'|=============================================================================

'.-----------------------------------------------------------------------------
'| JzReadINI : Reads a Value String from JzINIFile,JzSection,JzEntry
'|    Returns False if unsuccessfull, all results fields with no changes
'|    Returns True if successfull:
'|        JzValuINI becames with phisical Value String (all after first "=")
'|        If JzValuINI contains at least a ";" separator, then
'|           JzValue becames as a REAL Value String (all before last ";")
'|           JzRemarks becames as Comments String (all after last ";")
'|        Else
'|           JzValue becames as a REAL Value String (= JzValuINI)
'|           JzRemarks becames as a Null String
'`-----------------------------------------------------------------------------
Public Function JzReadINI() As Boolean
Dim aboo As Boolean
Dim JzDefValuINI As String
Dim JzDefValue As String
Dim JzDefRemarks As String
Dim JzRet As String ' temp returned data
Dim i As Integer
Dim N As Long
    JzDefValuINI = JzValuINI
    JzDefValue = JzValue
    JzDefRemarks = JzRemarks
    'pre fill returned area with nulls
    JzValuINI = String$(255, 0)
    aboo = False ' for eventual error occourrence
    On Error GoTo NoINI
    N = GetPrivateProfileString(JzSection, JzEntry, vbNullString, JzValuINI, 255, JzIniFile)
    'N returns total caracteres after (=)
    If N <> 0 Then 'entry found
        JzValuINI = Trim(Mid(JzValuINI, 1, N)) 'streaps nulls and leading spaces
        JzDisMount ' split into JzValue and JzRemarks
        If Len(JzValue) = 0 Then
           aboo = False
        Else
           aboo = True
        End If
    Else
        aboo = False
    End If
'epilog
    If aboo = False Then 'we must restablish values
       JzValuINI = JzDefValuINI
       JzValue = JzDefValue
       JzRemarks = JzDefRemarks
    End If
    On Error GoTo 0
NoINI:
    JzReadINI = aboo
End Function

'.-----------------------------------------------------------------------------
'| JzWriteINI: Create/Update JzValueINI at JzINIFile, JzSection, JzEntry
'| Priority contents is JzValue:
'|    If JzValue significative then
'|       If JzRemarks significative then
'|          automatic mountage JzValuINI = JzValue + ";" + JzRemarks
'|       Else
'|          JzValuINI = JzValue
'|       End If
'|    Else
'|       If JzValuINI significative then
'|          automatic dismountage JzValuINI into JzValue and JzRemarks
'|       Else
'|          Do Nothing (exit sub without i-o)
'|       End If
'|    End If
'`-----------------------------------------------------------------------------
Public Sub JzWriteINI()
    If Not JzAnyMissing Then
       WritePrivateProfileString JzSection, JzEntry, JzValuINI, JzIniFile
    End If
End Sub
'.-----------------------------------------------------------------------------
'| JZGetINI : Do JzReadINI If fails then Creates the designed Entry.
'|            Be CAREFULL to Prepare defaults:
'|               Or (JzValuINI with phisical Value String and JzValue Empty)
'|               Or (JzValuINI Empty and JzValue with Real Value String
'|                   and (JzRemarks = Comments Or JzRemarks Empty))
'`-----------------------------------------------------------------------------
Public Sub JzGetINI()
    'do nothing if any inconsistence
    If Not JzAnyMissing Then
       If Not JzReadINI Then
          JzWriteINI
       End If
    End If
End Sub

'.----------------- INTERNAL PROCEDURES -------------------
'JzValuINI will be JzValu + " ; " + JzReMarks
Private Sub JzMount()
   Dim R As String
   Dim v As String
   v = Trim(JzValue)
   R = Trim(JzRemarks)
   JzValuINI = vbNullString
   If Len(v) > 0 Then
      If Len(R) > 0 Then
         JzValuINI = v & " ; " & R
      Else
         JzValuINI = v
      End If
   End If
End Sub

'Split JzValuINI into JzValue and JzRemarks
Private Sub JzDisMount()
Dim i As Integer
    JzValuINI = Trim(JzValuINI) 'Streaps
    If Len(JzValuINI) = 0 Or JzValuINI = ";" Then ' empty case
       JzValue = vbNullString
       JzRemarks = vbNullString
    Else ' there is a significative value
       i = InStrRev(JzValuINI, ";", Len(JzValuINI))
       If i = 0 Then 'no remarks
          JzValue = Trim(JzValuINI)
          JzRemarks = vbNullString
       Else
          JzValue = Trim(Mid(JzValuINI, 1, i - 1))
          If i = Len(JzValuINI) Then ' (;) found in last position
             JzRemarks = vbNullString
          Else
             JzRemarks = Trim(Mid(JzValuINI, i + 1, Len(JzValuINI) - i))
          End If
       End If
    End If
End Sub

'Test elements integrity
Private Function JzAnyMissing() As Boolean
    If Len(Trim(JzValue)) = 0 Then
       JzDisMount
    End If
   'Preliminar empty test
   If Len(Trim(JzValue)) = 0 Or _
      Len(Trim(JzEntry)) = 0 Or _
      Len(Trim(JzSection)) = 0 Then
         JzAnyMissing = True
         Exit Function
   End If
   JzMount
   'Default name if no one
   If Len(Trim(JzIniFile)) = 0 Then
      JzIniFile = App.Path & "\" & App.Title & ".Ini"  ' case you don't set it before
   End If
   If Not JzIniFile = PrevINIFile Then
      PrevINIFile = JzIniFile
      'Sure Dir structure is ready
      SureDirs PrevINIFile
   End If
   JzAnyMissing = False
End Function

'Traditional File Existence function
Public Function JzFileExists(FSpec As String) As Boolean
    JzFileExists = Not (Dir(FSpec) = vbNullString)
End Function

'Creates, if it is necessary, each of Dir and SubDir in a Path String
Private Function SureDirs(CompletePath As String) As String
   Dim s As String
   On Error GoTo SureDirsErr
   s = SplitPath(CompletePath)
   If Dir(s, vbDirectory) = vbNullString Then
      s = SureDirs(s)
      MkDir s
   End If
   SureDirs = CompletePath
SureDirsErr:
   On Error GoTo 0
End Function
'used into SureDirs to make reentrances
Private Function SplitPath(sPath As String) As String
   SplitPath = Mid$(sPath, 1, InStrRev(sPath, "\", Len(sPath)) - 1)
End Function

'..................... End Code for INI FILES ................................

'.----------------------------------------------------------------------------
'| Code for PARAMETERS & PREFERENCES
'=============================================================================
' Parameters: Are Settings for Enable/Disable, Visible/Unvisible, some
'     indicators as UserName, Encrypted Passwords, Titles, Phrases, Routines
'     Options, etc, as configurated by a Administrator or Installer Support.
'
' Preferences: Are Settings as Options, Colors and/or Design styles, certain
'              pre-fixed Variables, etc, maying updated by a App User at any
'              moment.
'
' As code programming, they are identicals. Sugested you put Preferences into
' one of current App Form and create a Sub App for Admins mode to do changes
' in Parameters. The Template will be for all.
'=============================================================================
' See Sample at frmPrefer
'
' We are using:
' a) Checked/Unchecked Menu Itens; and/or
' b) Enabled/Disabled Menu Itens; and/or
' c) Pre Initialized Variables
'
' By similarity you can create, for exemple, Visibled/Invisibled Menu Itens,
' if you want it.
'
' See also, at frmPrefer_Load, a direct code
' to produce a Customized Caption - you can
' create anyone those based.
'
'.------------- INTERNALS -----------------------------
'Preparing to Check/Uncheck
Private Sub PreParCheck(mnuTopic As Object)
  JzSection = "Checked on Menu"
  JzEntry = mnuTopic.Name
  JzValuINI = vbNullString
  JzRemarks = mnuTopic.Caption
  If Not mnuTopic.Checked = True Then
     JzValue = "0"
  Else
     JzValue = "1"
  End If
End Sub

'Preparing to Enable/Disable
Private Sub PreParEnable(mnuTopic As Object)
  JzSection = "Enabled on Menu"
  JzEntry = mnuTopic.Name
  JzValuINI = vbNullString
  JzRemarks = mnuTopic.Caption
  If Not mnuTopic.Enabled = True Then
     JzValue = "0"
  Else
     JzValue = "1"
  End If
End Sub

'Preparing to Get/Set Variables value
Private Sub PreParVariable(mnuTopic As Object, strDefault As String)
  JzSection = "Variables on Menu"
  JzEntry = mnuTopic.Name
  JzValuINI = vbNullString
  If Len(Trim(strDefault)) > 0 Then
     JzValue = strDefault
  Else
     JzValue = JzVarFromMenu(mnuTopic)
  End If
  JzRemarks = mnuTopic.Caption ' before changes
End Sub

'Shows Menus with pre initialized Variables
Private Sub DispPar(mnuTopic As Object, preVar As String)
  Dim s As String
  Dim i As Integer
  i = InStr(mnuTopic.Caption, "=")
  s = Mid(mnuTopic.Caption, 1, i)
  mnuTopic.Caption = s & preVar
End Sub

'.------------------ PUBLICS --------------------------
'Get/Create Menu Checked entries (default at design time)
Public Sub JzParCheck(mnuTopic As Object)
  PreParCheck mnuTopic
  JzGetINI
  If JzValue = "1" Then
     mnuTopic.Checked = True
  Else
     mnuTopic.Checked = False
  End If
End Sub

'Update INI parameter with actual Check status from Menu
Public Sub JzCheckToINI(mnuTopic As Object)
  PreParCheck mnuTopic
  JzWriteINI
End Sub

'Get/Create Menu Enabled entries (default at design time)
Public Sub JzParEnable(mnuTopic As Object)
  PreParEnable mnuTopic
  JzGetINI
  If JzValue = "1" Then
     mnuTopic.Enabled = True
  Else
     mnuTopic.Enabled = False
  End If
End Sub

'Update INI parameter with actual Enabled status from Menu
Public Sub JzEnableToINI(mnuTopic As Object)
  PreParEnable mnuTopic
  JzWriteINI
End Sub

'Returns Menu INI Variable Parameter value
Public Function JzParVariable(mnuTopic As Object, Optional xDefault As String = vbNullString) As String
  PreParVariable mnuTopic, xDefault
  JzGetINI
  DispPar mnuTopic, JzValue
  JzParVariable = JzValue
End Function

'Update INI parameter with a Variable from Menu
Public Sub JzVarToINI(mnuTopic As Object, strNewVar As String)
  PreParVariable mnuTopic, strNewVar
  JzWriteINI
  DispPar mnuTopic, strNewVar
End Sub

'Get Variable stored in a Menu Item
Public Function JzVarFromMenu(mnuTopic As Object) As String
  Dim s As String
  Dim i As Integer
  i = InStr(mnuTopic.Caption, "=")
  If (i = 0 Or i = Len(mnuTopic.Caption)) Then
     JzVarFromMenu = vbNullString
  Else
     JzVarFromMenu = Trim(Mid(mnuTopic.Caption, i + 1, i))
  End If
End Function

'User updating Variable - auto menu and INI upd
'will receive that from a InputBox
'missing VarDescription to show assumes menu item Caption
Public Function JzNewVariable(mnuTopic As Object, Optional strVarDescription As String = vbNullString) As String
  Dim v As String ' previous var
  Dim u As String ' resulted var by default or as typed
  Dim t As String ' previous var or "<empty>" to be showed
  Dim s As String ' final var description
  Dim i As Integer
  s = Trim(strVarDescription)
  If Len(s) = 0 Then
     ' redundant code but easy
     i = InStr(mnuTopic.Caption, "=")
     If (i = 0 Or i = Len(mnuTopic.Caption)) Then
        v = vbNullString
     Else
        v = Trim(Mid(mnuTopic.Caption, i + 1, i))
        'get menu caption but streapping "&"
        s = Replace(Trim(Mid(mnuTopic.Caption, 1, i - 1)), "&", vbNullString)
     End If
  Else
     v = JzVarFromMenu(mnuTopic)
  End If
  If Len(Trim(v)) = 0 Then
     t = "<empty>"
  Else
     t = v
  End If
  u = Trim(InputBox(s, " Actual is " & t, v))
  If Len(u) > 0 Then
     JzVarToINI mnuTopic, u
  End If
  JzNewVariable = u
End Function

'..................... End Code for Parameters & Preferences  ................

'.----------------------------------------------------------------------------
'| Code for RECENT FILES STORED IN MENU process
'=============================================================================
'=============================================================================
' See Sample at frmRecents

Private Function JzFindName(ByVal FSpec As String) As String
Dim P As Long
    JzFindName = vbNullString
    If Len(FSpec) > 0 Then
       P = InStrRev(FSpec, "\")
       If P > 0 Then
          JzFindName = Mid$(FSpec, P + 1)
       Else
          P = InStrRev(FSpec, ":")
                  If P > 0 Then
             JzFindName = Mid$(FSpec, P + 1)
          End If
       End If
        End If
End Function

Public Function JzFindPath(ByVal FSpec As String) As String
Dim P As Long
    JzFindPath = vbNullString
    If Len(FSpec) > 0 Then
       P = InStrRev(FSpec, "\")
       If P > 0 Then
          JzFindPath = Mid$(FSpec, 1, P)
       Else
          P = InStrRev(FSpec, ":")
          If P > 0 Then
             JzFindPath = Mid$(FSpec, 1, P)
          End If
           End If
    End If
End Function

' This is a visual reduced form of filepaths as showed in Menu Itens
Public Function ShortenFileSpec(ByVal FSpec As String, _
                                ByVal l As Long) As String
Dim P     As String
Dim N     As String
Dim LName As Long
Dim LLeft As Long
Dim NDots As Long
    If Len(FSpec) <= l Then
        ShortenFileSpec = FSpec
    Else
        P = JzFindPath(FSpec)
        N = JzFindName(FSpec)
        LName = Len(N)
        LLeft = l - LName
        If LLeft = 0 Then       ' LName = L
            ShortenFileSpec = N
        ElseIf LLeft < 0 Then   ' rotulo de arquivo longo demais
            ShortenFileSpec = Mid$(N, 1, l \ 2 - 2) & ".." & Mid$(N, l \ 2, Len(N) - (l / 2) + 1)
        Else
            NDots = LLeft \ 2
            ShortenFileSpec = Mid$(P, 1, NDots) & String$(NDots, ".") & N
        End If
    End If
End Function

'secure returns filepath from buffer
Public Function RecentPath(Index As Long) As String
   If Not Index > UBound(RecentsArray) Then
      JzValuINI = RecentsArray(Index)
      JzDisMount
      RecentPath = JzValue
   Else
      RecentPath = vbNullString
   End If
End Function

'Validate and Put a JzValue to Recents Queue
'We need discard last Recent Path to Include a new one at first place at Menu
Private Sub PushMenuItens(Mnu As Object)
Dim k As Long
    If Mnu.Count < MaxRecentFiles Then  ' expand if not in maximum
        k = Mnu.UBound + 1
        Load Mnu(k)
        ReDim Preserve RecentsArray(k)
    End If
    ' We will disponibilize the First for a new Recent String
    For k = Mnu.UBound To 1 Step -1
        Mnu(k).Caption = Mnu(k - 1).Caption
        RecentsArray(k) = RecentsArray(k - 1)
    Next k
End Sub

Public Function JzToRecents(FileSpec As String, Mnu As Object) As Boolean
   Dim k As Long
   Dim aboo As Boolean
   JzRemarks = vbNullString
   aboo = False
   If Len(Trim(FileSpec)) > 0 Then ' significant
      If JzFileExists(FileSpec) Then  'and existent
             ' Let's see if it is already in Recent List
         For k = 0 To UBound(RecentsArray)
             If RecentPath(k) = FileSpec Then
                Exit For
             End If
         Next k
         If k > UBound(RecentsArray) Then ' it wasn't there so we will put it
            PushMenuItens Mnu ' roll down all in menu
            RecentsArray(0) = FileSpec     ' and store as first
            Mnu(0).Caption = ShortenFileSpec(FileSpec, ShortLen)
         End If
         aboo = True 'all ok
      Else
         'file no more exists
         Call MsgBox(FileSpec, vbCritical, "File No More Exists!")
      End If
   End If
   JzToRecents = aboo
End Function

' Load Menu Itens with Recent Files already in INI file
' Pre: JzINIFile contains Ini Specs for INI file path (at form load?)
'.------------ mnuRecentFiles(), ini & RecentsArray() ------------------
Public Sub JzGetRecentsFromINI(Mnu As Object)
Dim aboo As Boolean
Dim j As Long
Dim k As Long
Dim l As Long
    aboo = False
    j = 0
    k = 0
    JzSection = "Recent Files"
    On Error Resume Next
    Do
        JzEntry = CStr(k + 1) & "."
        If JzReadINI Then
           If Not JzValue = "_" Then ' We will consider as a Deleted Entry
              If JzFileExists(JzValue) Then  ' File still there
                 If j = 0 Then
                    aboo = True 'first valid entry
                 End If
                 RecentsArray(j) = JzValuINI
                 j = j + 1
                 ReDim Preserve RecentsArray(j)
              End If
           End If
           k = k + 1
        Else
           Exit Do
        End If
    Loop
    If aboo Then
       l = j - 1
       ReDim Preserve RecentsArray(l)
       ' first the most of recents
       Mnu(0).Caption = ShortenFileSpec(RecentsArray(0), ShortLen)
       ' now any more if they are there
       If j > 0 Then
          For j = 1 To l
              Load Mnu(j) ' create a new menu item for this
              Mnu(j).Caption = ShortenFileSpec(RecentsArray(j), ShortLen)
          Next j
       End If
    End If
    On Error GoTo 0
End Sub

'Write Actual Recent Files List to INI File
Public Sub JzToINIFile()
Dim k As Long
Dim l As Long
    l = UBound(RecentsArray)
    If l >= MaxRecentFiles Then
       l = MaxRecentFiles - 1
    End If
    JzSection = "Recent Files"
    For k = 0 To l
        JzValuINI = RecentsArray(k)
        JzDisMount
        JzEntry = CStr(k + 1) & "."
        JzWriteINI
    Next k
    'null value for next entries if was there
    Do
        JzEntry = CStr(k + 1) & "."
        If JzReadINI Then
           JzValuINI = "_"
           JzValue = vbNullString
           JzRemarks = vbNullString
           JzWriteINI
           k = k + 1
        Else
           Exit Do
        End If
    Loop
    On Error GoTo 0
End Sub

'When user clicks on mnuRecentFiles ...
Public Function JzFromRecents(Index As Integer) As String
   Dim l As Long
   Dim k As Long
   l = UBound(RecentsArray)
   k = Index
   If k > l Then
      k = l - 1
   End If
   JzValuINI = RecentsArray(k)
   JzDisMount
   If Not JzFileExists(JzValue) Then
      'file no more exists
      Call MsgBox(JzValue, vbCritical, "File No More Exists!")
      JzFromRecents = vbNullString
   Else
      JzFromRecents = JzValue
   End If
End Function

'..................... End Code for Recent Files .................................
