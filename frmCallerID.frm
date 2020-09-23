VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCallerID 
   Caption         =   "Caller ID Form"
   ClientHeight    =   4755
   ClientLeft      =   4035
   ClientTop       =   2805
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   5055
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM DD"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Text            =   "Date"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Text            =   "Time"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Text            =   "Phone Number"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Text            =   "Name"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.PictureBox picHook 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   2280
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Connect 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtStatus 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InputLen        =   1
   End
   Begin VB.Label Label6 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Phone Number"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Recieved CallerID Info:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Caption         =   "Disconnected"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCallerID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed By:  Paul Bergman
'
' Notes:          This program is meant to be a reference to creating
'                 your own caller ID program, not an off the shelf application
'                 that works for everyone.  There are several aspects
'                 of this program that will have to be taken into account
'                 before running it successfully.
'
'                 1.You must have a caller ID compatible modem
'                 2.You will need the correct modem command to turn on
'                   caller ID in your program
'                 3.You will need the correct CommPort
'                 4.You should learn the caller ID patterns/tags so you can
'                   modify the code to work for you.



Option Explicit

'Private Constants
Private Const chEventStart = "+"
Private Const DefDataPath = "C:\"

Private buffer As String   ' no use
Private strNumber As String
Private strName As String
Private strDate As String
Private strTime As String
Private onHook As Boolean
Private unwantedChar As Boolean
Private isName As Boolean
Private isPhone As Boolean


Private Sub Connect_Click()
On Error GoTo Connect_Click_Err

  If (Connect.Caption = "&Connect") Then                      ' This menu item will open or close the com port
    
    lblStatus.Caption = "Connected"
    
    If Not MSComm1.PortOpen Then                              ' Open the comm port if not already open
        MSComm1.PortOpen = True
    End If

    If Not MSComm1.PortOpen Then                              ' if there is a problem opening the port
        MsgBox "Cannot open comm port " & MSComm1.CommPort    ' display an error first
        End                                                   ' bail out of the program
    End If

    ' Initialize communications and update app UI
    MSComm1.RThreshold = 1                                    ' Generate a receive event on every character received
    MSComm1.InputLen = 1                                      ' Read the receive buffer 1 char at a time
    
    ' Make sure that you send the correct Modem Command
    MSComm1.Output = "AT+VCID=1" & vbCr                       ' Send command to put Identifier in event mode and receive serial number
    Connect.Caption = "Dis&connect"                         ' Change the menu to reflect opposite of port status

  Else
    MSComm1.PortOpen = False                                  ' Close the port and update app UI
    Connect.Caption = "&Connect"                            ' Change the menu to reflect opposite of port status
    lblStatus.Caption = "Disconnected"
    'txtStatus = ""
  End If
  
  Exit Sub

Connect_Click_Err:

  If Err.Number = 8005 Then
    MsgBox "Unable to connect to modem, port already open.", vbCritical, "Error"
  
  End If
End Sub

Private Sub Form_Load()
  onHook = False
  Connect_Click
End Sub


' Close the CommPort if it is still open when program closes
Private Sub Form_Unload(Cancel As Integer)
  If MSComm1.PortOpen Then
    MSComm1.PortOpen = False
  End If
End Sub


Private Sub MSComm1_OnComm()
  Static stEvent             As String                       'storage for an Identifier event
  Dim stComChar               As String * 1                   'temporary storage for received comm port data


  Select Case MSComm1.CommEvent

    Case comEvReceive                                       ' Received RThreshold # of chars.

      '----------------------------------------------------------------------------------------------
      'The following illustrates how the Identifier is designed
      'to make authoring software easy as '123' for developers:
      '1) Look for a "+" character which indicates the beginning of an event
      '2) Save subsequent characters until you detect a carriage return
      '3) Process the Event
      '----------------------------------------------------------------------------------------------
      Do
        stComChar = MSComm1.Input                         'read 1 character .Inputlen = 1
        txtStatus.Text = txtStatus.Text + stComChar
        
        If unwantedChar = False Then
          
          Select Case stComChar
  
              Case ""                                   'ascii character 16
                  onHook = Not onHook
                  whatColour
                  unwantedChar = True
              Case vbCr                                   'The CR indicates the end of the Identifier Event
                  ProcessEvent stEvent                    'Process the Identifier event
                  stEvent = ""
              Case Else
                  stEvent = stEvent + stComChar           'Save everything between the + and CR
          End Select
        Else
          unwantedChar = False
        End If
        Loop While MSComm1.InBufferCount                      'Loop until all characters in receive buffer are processed
      
  End Select
End Sub

Private Sub ProcessEvent(sTemp As String)
  Dim lc As Long

  If Len(sTemp) < 1 Then Exit Sub
  
  Select Case Mid(sTemp, 2, 4)
    ' resets variables to false to signify the start of a new call
    Case "RING"
      isName = False
      isPhone = False
    
    ' extracts the date from the DATE tag
    Case "DATE"
      txtDate = Left(Right(sTemp, 4), 2) & "/" & Right(sTemp, 2)
    
    ' extracts the time from the TIME tag
    Case "TIME"
      txtTime = Left(Right(sTemp, 4), 2) & ":" & Right(sTemp, 2)
    
    ' if NAME tag has CID Name Information, it will be stored in txtName
    ' for the user to see
    Case "NAME"
      If Len(sTemp) > 6 Then
        txtName = Right(sTemp, Len(sTemp) - 6)
        isName = True
      End If
    
    
    ' This case will probably not be used if you reside in USA, I needed it
    ' becase I live in Winnipeg, CA and phone systems send Phone# in MESG tag
    ' format which makes it more confusing.
    
    
    ' Some users will need to customize/Remove this tag pending on there
    ' geographical location
    Case "MESG"
      If Len(sTemp) > 20 Then
        'txtNumber = Right(sTemp, Len(sTemp) - 6)
        txtNumber = ""
        For lc = 18 To 5 Step -2
          txtNumber.Text = Mid(Right(sTemp, Len(sTemp) - 6), lc, 1) & txtNumber.Text
        Next lc
        isPhone = True
        If isName = False Then
          isName = True
          txtName.Text = "Unknown Name"
        End If
      ElseIf Len(sTemp) = 12 And isPhone = False Then
        txtNumber.Text = "Unknown Number"
        isPhone = True
      ElseIf Len(sTemp) = 12 And isName = False Then
        txtNumber.Text = "Unknown Name"
        isName = True
      End If
    
    
    ' Simply stores the number of the caller in txtNuber
    Case "NMBR"
      txtNumber = Right(sTemp, Len(sTemp) - 6)
  End Select
End Sub

' Potential Future use
' A possible way of finding out if the user answered the call or not
Private Sub whatColour()
  If onHook = False Then
    picHook.BackColor = vbRed
  Else
    picHook.BackColor = vbGreen
  End If

End Sub

'example of data being sent through the phonelines (for me)
'This is why I needed to use the MESG tag

'AT+VCID=1
'
'OK
'
'RING
'
'DATE=0215
'TIME=2324
'NAME=P BERGMAN
'MESG=030735353531323334   (phone # 555-1234)
'00
'MESG=0000
'MESG=00
'RING
'
'RING

