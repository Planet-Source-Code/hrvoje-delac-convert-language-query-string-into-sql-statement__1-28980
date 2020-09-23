VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "String -> SQLStmt"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Parse"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "condo +""by the sea"" +beach -expensive"
      ToolTipText     =   "condo +""by the sea"" +beach -expensive"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Keywords:       "
      Height          =   195
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "condo +""by the sea"" +beach -expensive"
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hrvoje Delac, hrvoje.delac@st.tel.hr, June 11th 2000

Private Sub Command1_Click()
    Text2 = Parse(Text1)
    Text1.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Function Parse(ByVal strString As String)
    '--------------------------------------------------------
    'function converts search string (like in Altavista) into
    'SQL query string for database search
    '
    'To make this work in your database, you need to
    'replace "table" and "field" with appropriate values
    '--------------------------------------------------------
    
    Dim intBlank As Integer       'first intBlank space position
    Dim intNextBlank As Integer   'Next intBlank space position (d)
    Dim intCount As Integer       'intCount variable
    Dim strFirstLeft  As String   'first character following intBlank
    Dim strSecondLeft As String   'first character following strFirstLeft
    Dim strSQLStmt As String      'SQL statement
    Dim strWord As String         'each Word within string
    Dim strPhrase As String       'Phrase within quotations
    Dim strChars As String        'All chars. Used for error checking.
    Dim blnAnyChars As Boolean    'Is there any alpha and num characters in strString
                                  'Used for error checking.
    
    Const FIELD As String = "field" ' replace this with actual field name you are searching
    Const TABLE As String = "table" ' replace this with actual table name you are searching
    
    'Begin Error checking
    strChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    blnAnyChars = False
    For intCount = 1 To 36
        If InStr(1, UCase(strString), Mid(strChars, intCount, 1)) <> 0 Then
            blnAnyChars = True
        End If
    Next intCount
    If Not blnAnyChars Then Exit Function
    'End Error checking
    
    strString = Chr(32) & Trim(strString) & Chr(32)
    intCount = 0
    intBlank = 0
    strSQLStmt = "SELECT * FROM " & TABLE & " WHERE"
      
    Do While InStr(strString, Chr(32)) <> 0
        intBlank = InStr(strString, Chr(32))
        
        If intBlank = 0 Then
            strFirstLeft = Mid(strString, intBlank, 1)
            strSecondLeft = Mid(strString, intBlank + 1, 1)
        Else
            strFirstLeft = Mid(strString, intBlank + 1, 1)
            strSecondLeft = Mid(strString, intBlank + 2, 1)
        End If
        
        intNextBlank = InStr(intBlank + 1, strString, Chr(32))
        
        If strFirstLeft = """" Then
            strWord = Mid(strString, InStr(intBlank, strString, Chr(34)) + 1, InStr(intBlank + 2, strString, Chr(34)) - 3)
            intNextBlank = InStr(intBlank + 2, strString, Chr(34)) + 1
        Else
            If strSecondLeft = """" Then
                strWord = Chr(32) & Chr(32) & Mid(strString, InStr(intBlank, strString, Chr(34)) + 1, InStr(intBlank + 4, strString, Chr(34)) - 4)
                intNextBlank = InStr(intBlank + 4, strString, Chr(34)) + 1
            Else
                strWord = Mid(strString, 1, InStr(intBlank + 2, strString, Chr(32)))
            End If
        End If

        Select Case strFirstLeft
            Case "+":
                If intCount <> 0 Then strSQLStmt = strSQLStmt & " AND"
                strSQLStmt = strSQLStmt & " " & FIELD & " LIKE '%"
                strSQLStmt = strSQLStmt & Trim(Mid(strWord, 3))
                strSQLStmt = strSQLStmt & "%'"
            Case "-":
                If intCount <> 0 Then strSQLStmt = strSQLStmt & " AND"
                strSQLStmt = strSQLStmt & " " & FIELD & " NOT LIKE '%"
                strSQLStmt = strSQLStmt & Trim(Mid(strWord, 3))
                strSQLStmt = strSQLStmt & "%'"
            Case Chr(32), "":
                strSQLStmt = strSQLStmt
            Case Is <> "+", "-", Chr(32):
                If intCount <> 0 Then strSQLStmt = strSQLStmt & " OR"
                strSQLStmt = strSQLStmt & " " & FIELD & " LIKE '%"
                strSQLStmt = strSQLStmt & Trim(strWord)
                strSQLStmt = strSQLStmt & "%'"
        End Select
        
        intCount = intCount + 1
        strString = Right(strString, Len(strString) - intNextBlank + 1)
        
        If strFirstLeft = "" Then
            Exit Do
        End If
    Loop
       
    Parse = strSQLStmt
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
        KeyAscii = 0
    End If
End Sub
