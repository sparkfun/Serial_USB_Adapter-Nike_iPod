VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nike+iPod Serial"
   ClientHeight    =   5715
   ClientLeft      =   4635
   ClientTop       =   4245
   ClientWidth     =   11055
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11055
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboBaud 
      Height          =   315
      ItemData        =   "Main.frx":0ECA
      Left            =   1320
      List            =   "Main.frx":0EDD
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdBreak 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9000
      Top             =   5160
   End
   Begin VB.TextBox txtComPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "2"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Begin Listening"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9960
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtResponse 
      Height          =   2055
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3625
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Main.frx":0F04
   End
   Begin RichTextLib.RichTextBox txtPods 
      Height          =   1215
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2143
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Main.frx":0F86
   End
   Begin RichTextLib.RichTextBox txtData 
      Height          =   1215
      Left            =   2760
      TabIndex        =   11
      Top             =   1560
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2143
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Main.frx":1008
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Comm:"
      Height          =   315
      Left            =   480
      TabIndex        =   16
      Top             =   240
      Width           =   795
   End
   Begin VB.Label lblPackets 
      Caption         =   "0"
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Packets Received:"
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Footpod data:"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Footpods heard:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Raw Data Received:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Speed:"
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   645
      Width           =   795
   End
   Begin VB.Shape sptConnect 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblStat 
      Caption         =   "Status:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu File_Menu 
      Caption         =   "&File"
      Begin VB.Menu Spacer 
         Caption         =   "------"
      End
      Begin VB.Menu File_Close 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About_Menu 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

' 2-2-07 Nathan Seidle Spark Fun Electronics
' Initial listening coding for the Nike+iPod serial adapter
' Sends the two commands to the receiver to start listening
' Form displays all received HEX code

'iPod receiver communicates at 57600bps 8-n-1

'There are two init strings that the iPod Nano sends to the receiver. These
'two HEX strings put the receiver into listening mode.
'Sent 1st : 0xFF550409070025C7
'Sent 2nd : 0xFF55020905F0

'The first two bytes FF 55 seem to be a universal header
'The last byte is a addition CRC - always 0x54 remainder?

Dim CommPort As Integer
Dim CommSpeed As Double

Dim Stop_Waiting As Boolean

Private Sub About_Menu_Click()
    frmAbout.Show
End Sub

Private Sub cboBaud_Click()
    CommSpeed = cboBaud.List(cboBaud.ListIndex)
End Sub

Private Sub cmdBreak_Click()
    Stop_Waiting = True
End Sub


Private Function Hex_Convert(nate As String) As Integer
'This function takes a two character string and converts it to an integer
Dim new_hex As Long
Dim temp As Integer
Dim i As Integer
    
    new_hex = 0
    
    For i = 0 To Len(nate) - 1
    
        'Peel off first letter
        temp = AscB(Right(Left(nate, 1 + i), 1))
        
        'Convert it to a number
        If temp >= AscB("A") And temp <= AscB("F") Then
            temp = temp - AscB("A") + 10
        ElseIf temp >= AscB("0") And temp <= AscB("9") Then
            temp = temp - AscB("0")
        End If
        
        'Shift the number
        new_hex = new_hex * 16 + temp
    
    Next
    
    Hex_Convert = new_hex

End Function

Private Sub cmdClear_Click()
    txtResponse.Text = ""
    txtPods.Text = ""
    txtData.Text = ""
End Sub

Private Sub cmdStart_Click()
        
On Error GoTo EH
    
    Dim heard_ids(100) As String
    Dim pod_id As String
    Dim pod_data As String
    
    cmdStart.Enabled = False
    cmdBreak.Enabled = True
    Stop_Waiting = False
    
    'If port already opened then close it
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    
    'Check to see if the PIC is outputting the load key ASC(5)
    lblStatus.Caption = "Opening Port"
    MSComm1.CommPort = CommPort
    MSComm1.InputLen = 1
    MSComm1.InBufferSize = 512
    MSComm1.Settings = "57600,n,8,1"
    MSComm1.PortOpen = True
    
    Dim temp As String
    Dim incoming_string As String
    Dim new_byte
    
    txtResponse.Text = ""
    txtPods.Text = ""
    txtData.Text = ""
    
    'Clear the buffer of heard pod IDs
    For i = 0 To 100
        heard_ids(i) = "0"
    Next i
    
    output_string = ""
    output_string = output_string & "FF550409070025C7" '1st init string
    
    Record_Length = Len(output_string)
    
    For j = 1 To (Record_Length / 2)
        temp = Right(Left(output_string, (j * 2)), 2)
        temp = Hex_Convert(temp)
            
        MSComm1.Output = Chr(temp)
    Next j
    
    'Now listen for 'FF 55 04 09 00 00 07 EC'
    i = 0
    incoming_string = ""
    Do
        If Stop_Waiting = True Then Exit Do
        
        If MSComm1.InBufferCount > 0 Then
            new_byte = MSComm1.Input
            
            incoming_string = incoming_string & new_byte
        
            i = i + 1
        
        End If
        
        If i = 8 Then Exit Do
        DoEvents
    Loop
    txtResponse.Text = txtResponse.Text & "1st: " & StringToHex(incoming_string)
    
    output_string = ""
    output_string = output_string & "FF55020905F0" '2nd init string
    Record_Length = Len(output_string)
    
    For j = 1 To (Record_Length / 2)
        temp = Right(Left(output_string, (j * 2)), 2)
        temp = Hex_Convert(temp)
            
        MSComm1.Output = Chr(temp)
    Next j
    
    'Now listen for 'FF 55 04 09 00 00 07 EC'
    incoming_string = ""
    i = 0
    Do
        If Stop_Waiting = True Then Exit Do
        
        If MSComm1.InBufferCount > 0 Then
            new_byte = MSComm1.Input
            incoming_string = incoming_string & new_byte
            i = i + 1
        End If
        
        If i = 8 Then Exit Do
        DoEvents
    Loop
    
    txtResponse.Text = txtResponse.Text & Chr$(13) & "2nd: " & StringToHex(incoming_string)

    lblStatus.Caption = "Listening"
    sptConnect.FillColor = &HFF00&
    Do
        If Stop_Waiting = True Then Exit Do
        
        'Now listen for pod responses
        incoming_string = ""
        i = 0
        Do
            If Stop_Waiting = True Then Exit Do
            
            If MSComm1.InBufferCount > 0 Then
                new_byte = MSComm1.Input
                incoming_string = incoming_string & new_byte
                i = i + 1
            End If
            
            If i = 34 Then Exit Do
            DoEvents
        Loop
        
        lblPackets.Caption = lblPackets.Caption + 1
        
        txtResponse.Text = txtResponse.Text & Chr$(13) & "Heard: " & StringToHex(incoming_string)
        
        'Parse Pod ID out of incoming string
        pod_id = Right(Left(incoming_string, 11), 4)
        pod_id = StringToHex(pod_id)
        
        'Update list of heard pods
        For i = 0 To 100
            'Search through ids to see if we already have this one
            If heard_ids(i) = pod_id Then Exit For
                
            
            If heard_ids(i) = "0" Then
                'If new then add it to the list
                heard_ids(i) = pod_id
                Exit For 'End of the list indicator
            End If
        Next i
        
        txtPods.Text = ""
        For i = 0 To 100
            If heard_ids(i) = "0" Then Exit For
            
            If i = 0 Then
                txtPods.Text = "ID: " & heard_ids(i)
            Else
                txtPods.Text = txtPods.Text & Chr$(13) & "ID: " & heard_ids(i)
            End If
        Next i
        
        'Parse Pod data out of incoming string
        pod_data = Right(Left(incoming_string, 34), 23)
        pod_data = StringToHex(pod_data)
        txtData.Text = txtData.Text & Chr$(13) & "Data: " & pod_data
    Loop
    
    'We must have gotten to this point because a stop_waiting flag was set
    'Close the port
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    cmdBreak.Enabled = False
    cmdStart.Enabled = True

    Exit Sub

EH:
    
    MsgBox Err.Number & " = " & Err.Description
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    cmdBreak.Enabled = False
    cmdStart.Enabled = True
    
End Sub

Private Sub File_Close_Click()
    Unload Me
    End
End Sub

Private Sub SetDefaults()
    cboBaud.ListIndex = 3 '57600 bps
    CommPort = 1
End Sub

Private Sub Form_Load()

On Error GoTo EH

    'Read in the settings file - if available
    Open App.Path & "\settings.txt" For Input As #1
        Input #1, CommPort, CommSpeed
        
    Close #1
    
    If CommPort = 0 Then CommPort = 1
    
    txtComPort.Text = CommPort
    
    If CommSpeed = 9600 Then cboBaud.ListIndex = 0
    If CommSpeed = 19200 Then cboBaud.ListIndex = 1
    If CommSpeed = 38400 Then cboBaud.ListIndex = 2
    If CommSpeed = 57600 Then cboBaud.ListIndex = 3
    If CommSpeed = 115200 Then cboBaud.ListIndex = 4
    
    sptConnect.BackColor = &HFF&
    lblPackets.Caption = 0
    
    Exit Sub

EH:
    
    If Err.Number = 53 Then 'No File found
        'Default values
        SetDefaults
    ElseIf Err.Number = 62 Then
        'Problem with the file, close it and set default options
        Close #1
        'Default values
        SetDefaults
    Else
        MsgBox Err.Number & " = " & Err.Description
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo EH
    
    Stop_Waiting = True
    DoEvents
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    DoEvents
    
    Open App.Path & "\settings.txt" For Output As #1
        Write #1, CommPort, CommSpeed
        
    Close #1
    
    DoEvents

    Unload Me
    End
    
EH:
    MsgBox Err.Number & " = " & Err.Description
    
End Sub

Private Function StringToHex(s As String) As String
        Dim i As Integer
        Dim j As Integer
        Dim result As String
        
    For i = 1 To Len(s)
        j = Asc(Mid(s, i, 1))
        result = result & " " & IIf(j < 17, "0" & Hex(j), Hex(j))
    Next i
    
    StringToHex = result
    
End Function

Private Function StringToDec(s As String) As String
        Dim result As String
        
    For i = 1 To Len(s)
        result = result & " " & Format(Asc(Mid(s, i, 1)), "000")
    Next i
    
    StringToDec = result
    
End Function

Private Sub GetSerial()
        Dim buffer As String
    
    If MSComm1.PortOpen = False Then Exit Sub
    
    If TermMode Then
        If MSComm1.InBufferCount > 0 Then
            buffer = MSComm1.Input
            Select Case TermType
                Case "asc"
                
                Case "dec"
                    buffer = StringToDec(buffer)
                Case "hex"
                    buffer = StringToHex(buffer)
                Case Else
                
            End Select
            
            With txtTerm
                If Len(.Text) > 15000 Then
                    .Text = Right(.Text, 10000)
                End If
                .SelStart = Len(.Text)
                .SelText = buffer
                .SelStart = Len(.Text)
            End With
        End If
    End If
    
End Sub

Private Sub Timer1_Timer()
    DoEvents
    GetSerial
    DoEvents
End Sub

Private Sub txtComPort_Change()
    
On Error GoTo Error_Handler
    
    If CInt(txtComPort.Text) > 99 Then txtComPort.Text = "99"

    CommPort = CInt(txtComPort.Text)
    
    Exit Sub
    
Error_Handler:
    MsgBox "Comport number must be a number from 1 to 99."
    txtComPort.Text = "1"

End Sub

