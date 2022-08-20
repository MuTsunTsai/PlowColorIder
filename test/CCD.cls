VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit


Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Const LF_FACESIZE = 32

Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Public Enum EOrientation
    eoPortrait = 1
    eoLandscape
End Enum

Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF
    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC
    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C
    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2001
    CFERR_MAXLESSTHANMIN = &H2002
    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003
    CCERR_CHOOSECOLORCODES = &H5000
End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Private m_lApiReturn As Long
Private m_lExtendedError As Long

Public Property Get APIReturn() As Long
    APIReturn = m_lApiReturn
End Property

Public Property Get ExtendedError() As Long
    ExtendedError = m_lExtendedError
End Property

Function VBGetOpenFileName(Filename As String, Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, Optional InitDir As String, _
                           Optional DlgTitle As String, Optional DefaultExt As String, _
                           Optional Owner As Long = -1, Optional flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, S As String, afFlags As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0

    With opfile
        .lStructSize = Len(opfile)
        
        .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
                (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
                 (-ReadOnly * OFN_READONLY) Or _
                 (-HideReadOnly * OFN_HIDEREADONLY) Or _
                 (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                      OFN_ENABLETEMPLATE)))

        If Owner <> -1 Then .hWndOwner = Owner    ' InitDir can take initial directory string
        .lpstrInitialDir = InitDir                ' DefaultExt can take default extension
        .lpstrDefExt = DefaultExt                 ' DlgTitle can take dialog box title
        .lpstrTitle = DlgTitle
        
        Dim ch As String, I As Integer
        For I = 1 To Len(Filter)
            ch = Mid$(Filter, I, 1)
            If ch = "|" Or ch = ":" Then
                S = S & vbNullChar
            Else
                S = S & ch
            End If
        Next
        S = S & vbNullChar & vbNullChar
        .lpstrFilter = S
        .nFilterIndex = FilterIndex
            
        S = Filename & String$(MAX_PATH - Len(Filename), 0)
        .lpstrFile = S
        .nMaxFile = MAX_PATH
        S = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
        .lpstrFileTitle = S
        .nMaxFileTitle = MAX_FILE
        
        m_lApiReturn = GetOpenFileName(opfile)
        Select Case m_lApiReturn
        Case 1
            VBGetOpenFileName = True
            Filename = StrZToStr(.lpstrFile)
            FileTitle = StrZToStr(.lpstrFileTitle)
            flags = .flags
            FilterIndex = .nFilterIndex
            Filter = FilterLookup(.lpstrFilter, FilterIndex)
            If (.flags And OFN_READONLY) Then ReadOnly = True
        Case 0
            VBGetOpenFileName = False
            Filename = ""
            FileTitle = ""
            flags = 0
            FilterIndex = -1
            Filter = ""
        Case Else
            m_lExtendedError = CommDlgExtendedError()
            VBGetOpenFileName = False
            Filename = ""
            FileTitle = ""
            flags = 0
            FilterIndex = -1
            Filter = ""
        End Select
    End With
End Function

Private Function StrZToStr(S As String) As String
    StrZToStr = Left$(S, lstrlen(S))
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, S As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            S = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            S = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = S
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function





