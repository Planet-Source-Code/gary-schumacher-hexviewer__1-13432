VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form HexFileViewer 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Hex File View"
   ClientHeight    =   6828
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   8268
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6828
   ScaleWidth      =   8268
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   300
      Left            =   2640
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   2652
      _ExtentX        =   4678
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6528
      Width           =   8268
      _ExtentX        =   14584
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3493
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3493
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3493
            TextSave        =   "8:41 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3493
            TextSave        =   "12/13/00"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTFBox 
      Height          =   5172
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   7932
      _ExtentX        =   13991
      _ExtentY        =   9123
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"HexFileViewer.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   1080
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Label lblColumnHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Hex Off  00 01 02 03 04 05 06 07   08 09 0A 0B 0C 0D 0E 0F     Ascii Values"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   288
      Width           =   9360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "HexFileViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_strSourceFile As String          'Will contain entire file
Dim m_arrXlate(0 To 255, 2) As String * 2  'ascii to hex translation table



Private Sub Form_Load()
    
    AdjustControlLocations Me   'Center RichTextBox on Screen
    
    '
    'Build Translation table...ascii to hex
    'table values are 00 - FF...need to be 2 digits
    'Also, store "00" for an ascii non-printable char
    '....store "FF" for a printable char
    'printable chars are:
    '48 to 57   Ascii 0 - 9 -- numbers
    '65 to 90   Ascii A - Z -- Uppercase
    '97 to 122  Ascii a - Z -- Lowercase
    '
    'BuildXlateTable
    
    
End Sub
Private Sub BuildXlateTable()
'Now handled in C Dll...GS

'    Dim x As Integer                'Loop Counter
'    '
'    'Build Translation table...ascii to hex
'    'table values are 00 - FF...need to be 2 digits
'    'Also, store "00" for an ascii non-printable char
'    '....store "FF" for a printable char
'    'printable chars are:
'    '48 to 57   Ascii 0 - 9 -- Numbers
'    '65 to 90   Ascii A - Z -- Uppercase
'    '97 to 122  Ascii a - Z -- Lowercase
'    '
'    '
'    For x = 0 To 255
'        If x > &HF Then     'Double digits &h10-$hff need no adjustment
'            m_arrXlate(x, 1) = Hex$(x)
'        Else                'Add 0 to 0-F values
'            m_arrXlate(x, 1) = "0" & Hex$(x)
'        End If
'    Next x
'
'    '
'    'Now create ascii non-printable ("00") & printable ("FF")
'    '
'    For x = 0 To 255
'        Select Case x
'            Case 48 To 57   'Ascii 0 to 9
'                m_arrXlate(x, 2) = "FF"
'            Case 65 To 90   'Ascii A to Z
'                m_arrXlate(x, 2) = "FF"
'            Case 97 To 122  'Ascii a to z
'                m_arrXlate(x, 2) = "FF"
'            Case Else       'Nonprintable...
'                m_arrXlate(x, 2) = "00"
'        End Select
'     Next x
End Sub


Private Sub Form_Resize()
    Me.Refresh
    AdjustControlLocations Me   'Center RichTextBox on Screen
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me   'Unload Main Form
End Sub

Private Sub mnuExit_Click()
    Unload Me   'Unload Main Form
End Sub


Private Sub mnuOpen_Click()
        Dim FileNumber As Integer   'Used for File Operations
        Dim lngFileLen As Long      'Size of File in bytes
        Dim strMsg As String        'Error Message String
        
        On Error GoTo Error_mnuOpen
        
        
        CommonDialog1.CancelError = True
        CommonDialog1.Filter = "Any File (*.*)|*.*"
        CommonDialog1.ShowOpen
        Me.Caption = CommonDialog1.FileName
        '
        'Open Selected File...Store entire file in m_strSourceFile
        '
        m_strSourceFile = ""
        RTFBox.Text = ""            'Clear Text box
        RTFBox.Refresh              '...
        
        FileNumber = FreeFile
        Open CommonDialog1.FileName For Binary Access Read As FileNumber
            lngFileLen = LOF(FileNumber)    'Size of File
            m_strSourceFile = Space$(lngFileLen + 16) 'Allocate space to hold entire
                                                    '..File in string...Plus any remainder
            
            Get #FileNumber, , m_strSourceFile        'Load entire file into m_strSourceFile
           
        Close #FileNumber
        '
        'Build Output from File (m_strSourceFile)
        '
        Screen.MousePointer = vbHourglass  'Working....
        
        BuildFormattedOutput lngFileLen
        '
        'File is Displayed in Hex Format
        'Do a Little Cleanup
        '
        'ProgressBar.Value = 0              'Clear Progress Bar
        'ProgressBar.Visible = False        'Hide bar
        StatusBar.Panels(1).Text = ""      'Clear Statusbar
        Screen.MousePointer = vbDefault    'Back to mouse pointer

        Exit Sub                            'Bypass Error Routine
Error_mnuOpen:
        strMsg = "Error # " & Str(Err.Number) & " was generated in " _
          & "Open Routine" & vbCrLf
        strMsg = strMsg & Err.Description
        MsgBox strMsg, , "Error Encountered"
        Screen.MousePointer = vbDefault 'Back to mouse pointer
End Sub
Private Sub BuildFormattedOutput(lngFileLen As Long)
        Dim lngFivePercent As Long      'Used for progress bar...5% of total recs in file
        Dim lngRecords As Long          '# of complete 16 byte recs in file
        Dim intRemainder As Integer     '# of bytes remaining (none if evenly divisable by 16
        Dim strDummyOutput As String    'Dummy Formatted Output Record
        Dim lngSourceOffset As Long     'ptr into m_strSourceFile
        Dim strOutput() As String       'Array to hold Formatted Output
        Dim lngPtrOutput As Long        'ptr to each 79 byte record in strOutput
        
        Dim intPtrSrce As Integer       'ptr to each byte of current source record
                                        '...used with lngSourceOffset
        Dim strWork As String           'Temp work area for current formatted record
        Dim bytAscii As Byte            'ascii value of source record
        Dim strTemp As String           '2 byte hex value from Translate table (m_arrXlate)
        Dim strFormattedOutput As String    'Will contain entire formatted file
        Dim strValid As String              'Call from Xlate table "00" non-printable "FF" printable
        Dim strMsg As String               'Error Message String
        
        On Error GoTo Error_BuildFormattedOutput
        
        Me.Refresh                      'Redraw Form before going into Dll...
                                        'Lengthy processing in the Dll made
                                        'the screen look unsightly (sometimes)
        '
        'Calculate # of 16 byte records in file
        'Clng operation rounds #'s...so we must adjust later
        'So that last line of file is shown correctly
        '
        lngRecords = CLng(lngFileLen / 16)  '# of 16 byte recs in file
        intRemainder = lngFileLen Mod 16    'Plus any remaining bytes
        If intRemainder <> 0 Then
            '
            'Because of rounding, we may need to bump record count
            '
            If lngRecords * 16 < lngFileLen Then    'Q. Do we have enough records?
                lngRecords = lngRecords + 1         'A. No we need one more
            End If
            m_strSourceFile = m_strSourceFile & String(16 - intRemainder, Chr$(0)) 'Pad last record to equal 16 bytes
        End If
        '
        'Overlay Progress Bar on 2nd Panel of Status Bar
        'Use IF!!! File is large enough...
        'If lngRecords > 10000 Then          'Arbitrary Size
        '    SetupProgressBar
        '    ProgressBar.Max = lngRecords       '# of 16 byte Recs in File
        '    lngFivePercent = lngRecords * 0.05 'Calc 5% of records for progress bar
        'End If
        '
        'Load Output File With Dummy Records
        '                           1         2         3         4         5         6         7
        '                 0123456789012345678901234567890123456789012345678901234567890123456789012345678
        'strDummyOutput = "00000000: 00 00 00 00 00 00 00 00 - 00 00 00 00 00 00 00 00   ................ "
        '                 ^Hex off  ^ Offset 1st Hex value    ^ Offset 8th Hex value    ^Offset of 1st ascii value
        
        'ReDim strOutput(lngRecords) As String    'Make room for all 16 byte records in file
        'lngSourceOffset = 1                         'Point to current m_strSourceFile record
         
         strFormattedOutput = " "                'Initialize string
         strFormattedOutput = Space$(lngRecords * 79)  'Allocate space to hold entire file
         '
         'Format Each Record for Display
         '
         'For lngPtrOutput = 1 To lngRecords
         '   If lngRecords > 10000 Then              'Don't process if 'small' file
         '       If lngPtrOutput Mod lngFivePercent = 0 Then
         '           ProgressBar.Value = lngPtrOutput   'Show Progress in 5% increments
         '       End If
         '   End If
         '   strWork = strDummyOutput                'Copy dummy record template to work area
            'For intPtrSrce = 0 To 15                'Point to each location in source record (m_strSourceFile)
            '
            'Update Status Bar
            '
            StatusBar.Panels(1).Text = "Formatting File..."
            StatusBar.Refresh          'Keeps Statusbar text from disappearing
            '
            'Call Our Dll...The Entire (Except for any Remainder record) is now
            'being formatted in the Dll.  3MB files that used to take a couple of
            'minutes...now take about 10 seconds
            '
            HexDump m_strSourceFile, strFormattedOutput, lngFileLen, lngRecords
'                bytAscii = Asc(Mid(m_strSourceFile, lngSourceOffset + intPtrSrce, 1)) 'Extract char from source file
'                strTemp = m_arrXlate(bytAscii, 1)   'Convert to 2 byte hex value from table
'                strValid = m_arrXlate(bytAscii, 2)  'Will be "00" non-printable or "FF" printable
'                If intPtrSrce <= 7 Then             'Working on Left side of Output
'                    Mid(strWork, LeftHexOff + intPtrSrce * 3, 2) = strTemp
'                Else                                'Right side Hex #'s
'                    Mid(strWork, RightHexOff + (intPtrSrce - 8) * 3, 2) = strTemp
'                End If
'                If strValid = "FF" Then             '0-9, A-Z, a-z char found?
'                    Mid(strWork, AsciiOff + intPtrSrce, 1) = Chr$(bytAscii)
'                End If
            'Next intPtrSrce
            'strTemp = Hex$(lngPtrOutput - 1)        'Offset into File-1 (Left side of output)
            'Mid(strWork, 8 - Len(strTemp), Len(strTemp)) = strTemp
            'strOutput(lngPtrOutput) = strWork       'Save completed Formatted hex rec into array
            'lngSourceOffset = lngSourceOffset + 16  'Point to next 16 byte rec in m_strSourceFile
         'Next lngPtrOutput
        '
        'See if we had any remainder...i.e. last record has less than 16 bytes
        'If so...fill empty values with "  " blanks
        '
'        lngPtrOutput = lngPtrOutput - 1             'Point to last record in array ('Next' had incremented)
        'strWork = Right(strFormattedOutput, 79)
        '
        'Eliminate padded chr$(00) from last record (If not evenly divisable by 16)
        '
        If intRemainder <> 0 Then                   'Remainder Left? Skip if not
            For intPtrSrce = intRemainder To 15
                If intPtrSrce <= 7 Then             'Working on Left side of Output
                    Mid(strFormattedOutput, Len(strFormattedOutput) - 79 + LeftHexOff + intPtrSrce * 3, 2) = "  " 'Clear Location
                Else                                'Right side Hex #'s
                    Mid(strFormattedOutput, Len(strFormattedOutput) - 79 + RightHexOff + (intPtrSrce - 8) * 3, 2) = "  "
                End If
                Mid(strFormattedOutput, Len(strFormattedOutput) - 79 + AsciiOff + intPtrSrce, 1) = " "  'Clear unneeded chars on ascii side
            Next intPtrSrce
            If intRemainder <= 8 Then
                    Mid(strFormattedOutput, Len(strFormattedOutput) - 79 + DashOff, 1) = " " 'Clear Dash in output
            End If
            'Mid(strFormattedOutput, Len(strFormattedOutput) - 79, 79) = strWork 'Save updated Formatted hex rec
         End If
        '
        'Prepare for output to RichTextBox
        'Use Join to copy strOutput array into a string...strFormattedOutput
        'strFormattedOutput = ""                     'Initialize string
        'strFormattedOutput = Space$(lngRecords * 79)  'Allocate space to hold entire
                                                                        '..File in string
        'strFormattedOutput = Join(strOutput, "")    'Store array into string..."" means no delimeter
        RTFBox.Text = strFormattedOutput      'Display File in RichTextBox
        m_strSourceFile = ""                  'Free Some Memory
        Exit Sub                              'Bypass Error Routine
Error_BuildFormattedOutput:
        strMsg = "Error # " & Str(Err.Number) & " was generated in " _
          & "BuildFormattedOutput Routine" & vbCrLf
        strMsg = strMsg & Err.Description
        MsgBox strMsg, , "Error Encountered"
        'ProgressBar.Visible = False     'Hide bar
        Screen.MousePointer = vbDefault 'Back to mouse pointer
End Sub

Private Sub SetupProgressBar()
         '
         'Align Progress Bar to fit within Panel 2 of Status Bar
         '
'         With StatusBar.Panels(2)
'                ProgressBar.Left = .Left
'                ProgressBar.Width = .Width
'                ProgressBar.Height = StatusBar.Height - 40
'                ProgressBar.Top = StatusBar.Top + 20
'                ProgressBar.Visible = True
'             End With
'         StatusBar.Panels(1).Text = "Formatting File..."
'         StatusBar.Refresh          'Keeps Statusbar text from disappearing
End Sub

