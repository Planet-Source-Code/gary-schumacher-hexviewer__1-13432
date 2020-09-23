Attribute VB_Name = "Module1"
Option Explicit
  Declare Sub HexDump Lib _
  "C:\GAS\VB 60 Data\Hex File Viewer Submitted to Planet Source\HexView\Debug\HexView.dll" _
    (ByVal Source As String, ByVal Destination As String, _
    ByVal FileLen As Long, ByVal NumRecs As Long)
 '
 'Constants for Formatted Output Display
 '
 Public Const LeftHexOff = 11              'Offset of 1st Hex char in strWork (Left side)
 Public Const RightHexOff = 37             'Offset of 8th Hex char in strWork (Right side
 Public Const AsciiOff = 63                'Offset of 1st Ascii value in strWork
 Public Const DashOff = 35                 'Offset of Dash (between 8th & 9th Hex Value)
        '
Public Sub AdjustControlLocations(frm As Form)
    Dim strMsg As String
    Dim intCalcLeft As Integer      'use to calculate RichTextBox/ColumnHeader
    Dim intCalcHeight As Integer    '...positions on screen
    Dim strTemp As String           '
    Dim intLeft As Integer          'use to calculate progress/status bars
    Dim intWidth As Integer         '...
    '
    'Position RichTextBox & ColumnHeader label on form
    '
    With frm.RTFBox
        intCalcLeft = (frm.Width - .Width) / 2    'Calculate Center of RTextBox
        .Left = intCalcLeft                                 'Update RTextBox
        frm.lblColumnHeader.Left = intCalcLeft + 25         'Calc postion of Header Label
        
        intCalcHeight = (frm.Height - (.Top + 500)) - (frm.StatusBar.Height + 400)
        .Height = intCalcHeight                             'Update RTextBox
        '
        'Calculate width of RichTextBox
        'Otherwise Formatted Output will not be Displayed correctly
        '
        strTemp = String(79, "0")   '79 is width of formatted output record
        .RightMargin = frm.TextWidth(strTemp)   'Set right margin in RichTextBox
     End With
    
End Sub





