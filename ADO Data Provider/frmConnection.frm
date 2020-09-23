VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Connection..."
   ClientHeight    =   3390
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   2880
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Details"
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   3810
         TabIndex        =   15
         Top             =   1410
         Width           =   375
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         Top             =   1050
         Width           =   3015
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   1380
         Width           =   3015
      End
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "frmConnection.frx":000C
         Left            =   1200
         List            =   "frmConnection.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3000
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1200
         TabIndex        =   11
         Top             =   2055
         Width           =   3015
      End
      Begin VB.ComboBox cboDrivers 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1710
         Width           =   3015
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "DSN:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   0
         Top             =   405
         Width           =   360
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Userid:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   765
         Width           =   510
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1095
         Width           =   750
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Database:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1425
         Width           =   750
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Driver:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2100
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
    Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
    Const SQL_SUCCESS As Long = 0
    Const SQL_FETCH_NEXT As Long = 1


Private Sub cmdBrowse_Click()
On Error Resume Next

    With CommonDialog
        .CancelError = False
        .DialogTitle = "Browse..."
        If (InStr(1, LCase(cboDSNList.Text), "excel")) > 0 Then
            .Filter = "Microsoft Excel Database (*.xls)|*.xls"
        ElseIf (InStr(1, LCase(cboDSNList.Text), "access")) > 0 Then
            .Filter = "Microsoft Access Database (*.mdb)|*.mdb"
        Else
            .Filter = "All Databases (*.*)|*.*"
        End If
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        If Len(.FileName) > 0 Then
            txtDatabase.Text = .FileName
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandler

    Dim sConnect    As String
    Dim sADOConnect As String
    Dim sDSN        As String
       
    If (cboDSNList.ListIndex = 0) And (cboDrivers.Text = vbNullString) Then
        MsgBox "Mandantory field 'Driver' is not filled!", vbExclamation, "Warning!"
        lblStep3(5).ForeColor = vbRed
        cboDrivers.SetFocus
        Exit Sub
    End If
    
    If Len(txtDatabase.Text) = 0 Then
        MsgBox "Mandantory field 'Database' is not filled!", vbExclamation, "Warning!"
        lblStep3(4).ForeColor = vbRed
        txtDatabase.SetFocus
        Exit Sub
    End If
    
    If cboDSNList.ListIndex > 0 Then
        sDSN = "DSN=" & cboDSNList.Text & ";"
    Else
        sConnect = sConnect & "Driver={" & cboDrivers.Text & "};"
        sConnect = sConnect & "Server=" & txtServer.Text & ";"
    End If
    
    sConnect = sConnect & "Uid=" & txtUID.Text & ";"
    sConnect = sConnect & "Pwd=" & txtPWD.Text & ";"
    sConnect = sConnect & "Dbq=" & txtDatabase.Text & ";"
    
    sADOConnect = Chr(34) & "Provider=MSDASQL;" & sDSN & sConnect & Chr(34)


    With frmMain
        
        ' Aizveram aktîvo savienojumu, ja tas ir atvçrts:
        If .Connected Then
        
            .my_Connection.Close
            
            If (Err.Number = 0) Then
                .Connected = False
                .rtbSQL.BackColor = &H8000000F
                .WriteLog "Active connection closed"
            Else
                .WriteLog "Could not close active connection!"
                MsgBox "Could not close active connection!", vbCritical, "Error!"
            End If
            
        End If
        
        
        ' Atveram jaunu savienojumu:
        .WriteLog "Opening new connection..."
        .WriteLog sADOConnect
        
        '.my_Connection.ConnectionString = sADOConnect
        .my_Connection.Open sADOConnect
        
        If (Err.Number = 0) Then
            .Connected = True
            .Caption = txtDatabase & " - Data Provider"
            .rtbSQL.Enabled = .Connected
            .rtbSQL.BackColor = &H80000005
            .WriteLog "New connection successfully opened"
        Else
            .WriteLog "Could not open new connection!"
            MsgBox "Could not open new connection!", vbCritical, "Kïûda!"
        End If
        
    End With
    
    Unload Me
    
    Exit Sub
    
ErrHandler:
    
    MsgBox "Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Kïûda Atverot Jaunu Savienojumu!"
    
    frmMain.WriteLog "Error #" & Err.Number & ": " & Err.Description
    
End Sub

Private Sub Form_Load()
    GetDSNsAndDrivers
End Sub

Private Sub cboDSNList_Click()
On Error Resume Next

    ' Ja izvçlçts DNS, tad servera un dziòu lauki nav labojami!!!

    If cboDSNList.Text = "(None)" Then
        txtServer.Enabled = True
        cboDrivers.Enabled = True
    Else
        txtServer.Enabled = False
        cboDrivers.Enabled = False
    End If
    
    lblStep3(5).Enabled = cboDrivers.Enabled
    lblStep3(6).Enabled = txtServer.Enabled
    
End Sub

Sub GetDSNsAndDrivers()
On Error Resume Next

    ' Ðis kods tika automâtiski uzìenerçts un atgrieþ DSN un dziòu sarakstu

    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         ' Handle to the environment

    On Error Resume Next
    cboDSNList.AddItem "(None)"

    ' Get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
                cboDrivers.AddItem sDRV
            End If
        Loop
    End If
    
    ' Remove the dupes
    If cboDSNList.ListCount > 0 Then
        With cboDrivers
            If .ListCount > 1 Then
                i = 0
                While i < .ListCount
                    If .List(i) = .List(i + 1) Then
                        .RemoveItem (i)
                    Else
                        i = i + 1
                    End If
                Wend
            End If
        End With
    End If
    cboDSNList.ListIndex = 0
    
End Sub

