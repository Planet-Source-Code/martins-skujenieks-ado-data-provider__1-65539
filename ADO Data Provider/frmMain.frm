VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Provider"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   669
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbSQLBuffer 
      Height          =   855
      Left            =   8520
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSDataGridLib.DataGrid dgTable 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1062
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1062
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLog 
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5760
      Width           =   9975
   End
   Begin RichTextLib.RichTextBox rtbSQL 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   0   'False
      TextRTF         =   $"frmMain.frx":007D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADODC 
      Height          =   375
      Left            =   8160
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "Connection"
      Begin VB.Menu mnuConnectionNew 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuConnectionClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuSQL 
      Caption         =   "SQL"
      Begin VB.Menu mnuSQLExecute 
         Caption         =   "Execute"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Public Connected As Boolean
    Public my_Connection As ADODB.Connection
  
Private Sub Form_Load()
On Error Resume Next
  
    Width = 640 * Screen.TwipsPerPixelX
    Height = 480 * Screen.TwipsPerPixelY
    
    txtLog.Text = vbNullString
    
    With rtbSQL
        .Left = 0
        .Top = 0
        .Width = 634
        .Height = 299
    End With
    
    With dgTable
        .Left = 0
        .Top = 302
        .Width = 634
        .Height = 75
    End With
    
    With txtLog
        .Left = 0
        .Top = 380
        .Width = 634
        .Height = 46
    End With
    
    Show
    DoEvents
    
    Set my_Connection = New ADODB.Connection
    
    WriteLog "Ready"
    
    frmConnection.Show vbModal
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mnuConnectionClose_Click
End Sub

Private Sub mnuConnectionClose_Click()
On Error Resume Next
    If Connected = True Then
        WriteLog "Closing active connection..."
        my_Connection.Close
        WriteLog "Connection closed"
        Connected = False
        Caption = "Data Provider"
        rtbSQL.Enabled = Connected
        rtbSQL.BackColor = &H8000000F
    Else
        WriteLog "No open connection to close"
    End If
End Sub

Private Sub mnuConnectionNew_Click()
    frmConnection.Show vbModal
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "ADO Data Provider v" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & _
           "Developed by Martins Skujenieks" & vbCrLf & _
           "Last modified June 01, 2006 " & vbCrLf & vbCrLf & _
           "P.S. Please take a time to vote for this project! Thanks in advance!", , "About..."
End Sub

Private Sub mnuSQLExecute_Click()
On Error GoTo ErrHandler

    If Not Connected Then Exit Sub

    Dim sSQL As String
 
    If rtbSQL.SelLength > 0 Then
        sSQL = Mid(rtbSQL.Text, NVL(rtbSQL.SelStart, 1), rtbSQL.SelLength + 1)
    Else
        sSQL = rtbSQL.Text
    End If
    
    sSQL = Replace(Replace(sSQL, Chr(10), " "), Chr(13), " ")
    If Right(Trim(sSQL), 1) <> ";" Then sSQL = sSQL & ";"
    
    WriteLog "Executing SQL query..."
    WriteLog Replace(Replace(rtbSQL.Text, Chr(10), " "), Chr(13), "")
    
    '---------------------------------------------------------
    
    Dim my_Recordset As ADODB.Recordset
    Set my_Recordset = New ADODB.Recordset
    
    Set dgTable.DataSource = Nothing
    
    ' Set my_Recordset = my_Connection.Execute(rtbSQL.Text)
    my_Connection.CursorLocation = adUseClient
    my_Recordset.Open sSQL, my_Connection, adOpenDynamic, adLockOptimistic
    
    If (InStr(1, LCase(sSQL), "select") > 0) Then
        Set dgTable.DataSource = my_Recordset.DataSource
        WriteLog my_Recordset.RecordCount & " records returned"
    End If
    
    dgTable.SetFocus
    
    Exit Sub

ErrHandler:
    
    Select Case Err.Number
        Case -2147217865
            MsgBox "Table does not exist in DB!", vbCritical, "Error!"
        Case -2147217900
            MsgBox "Sintax error in SQL expression!", vbCritical, "Error!"
        Case Else
            MsgBox "Unknown error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Kïûda!"
    End Select
    
    WriteLog "Error #" & Err.Number & ": " & Err.Description
    
End Sub

Public Sub WriteLog(Text As String, Optional Color As Long = &H800000)
On Error Resume Next
    txtLog.ForeColor = Color
    txtLog.Text = txtLog.Text & vbCrLf & Str(Time) & "  " & Text
    If Len(Text) > 1 Then txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Function NVL(Value As Long, ValueIfNull As Long) As Long
On Error Resume Next
    If Value = 0 Then
        NVL = ValueIfNull
    Else
        NVL = Value
    End If
End Function

Private Sub rtbSQL_Change()
On Error Resume Next
        
    rtbSQL.Enabled = False
    rtbSQLBuffer.TextRTF = rtbSQL.TextRTF
    
    Dim p As Long
    Dim i As Long
    Dim l As Long
    Dim c As Long
    
    p = rtbSQL.SelStart
    
    For i = 1 To Len(rtbSQLBuffer.Text)
        
        l = 0
        c = vbBlack
        If UCase(Mid(rtbSQLBuffer.Text, i, 1)) = "*" Then l = 1: c = vbRed
        If UCase(Mid(rtbSQLBuffer.Text, i, 1)) = "(" Then l = 1: c = vbRed
        If UCase(Mid(rtbSQLBuffer.Text, i, 1)) = ")" Then l = 1: c = vbRed
        If UCase(Mid(rtbSQLBuffer.Text, i, 1)) = ";" Then l = 1: c = vbRed
        If UCase(Mid(rtbSQLBuffer.Text, i, 2)) = "AS" Then l = 2: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 6)) = "SELECT" Then l = 6: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 4)) = "FROM" Then l = 4: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 5)) = "WHERE" Then l = 5: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 8)) = "GROUP BY" Then l = 8: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 8)) = "ORDER BY" Then l = 8: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 6)) = "HAVING" Then l = 6: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 6)) = "UPDATE" Then l = 6: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 3)) = "SET" Then l = 3: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 6)) = "INSERT" Then l = 6: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 4)) = "INTO" Then l = 4: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 6)) = "VALUES" Then l = 6: c = vbBlue
        If UCase(Mid(rtbSQLBuffer.Text, i, 6)) = "DELETE" Then l = 6: c = vbBlue
        
        If (l > 0) Then
            rtbSQLBuffer.SelStart = i - 1
            rtbSQLBuffer.SelLength = l
            rtbSQLBuffer.SelColor = c
            rtbSQLBuffer.SelText = UCase(rtbSQLBuffer.SelText)
        End If
        
    Next
    
    rtbSQL.TextRTF = rtbSQLBuffer.TextRTF
    rtbSQL.Enabled = True
    rtbSQL.SetFocus
        
    rtbSQL.SelStart = p
    rtbSQL.SelLength = 0
    rtbSQL.SelColor = vbBlack

End Sub

Private Sub dgTable_GotFocus()
    Shrink 25
End Sub

Private Sub rtbSQL_GotFocus()
    Shrink 75
End Sub

Private Sub Shrink(Proportion As Single)
On Error Resume Next

    If Proportion > 80 Then Proportion = 80
    If Proportion < 20 Then Proportion = 20
    
    Proportion = Int(374 * Proportion / 100)
    
    rtbSQL.Height = Proportion
    dgTable.Top = Proportion + 3
    dgTable.Height = 374 - Proportion

End Sub
