VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   1560
      Top             =   3120
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Print Table Structure"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         TabIndex        =   7
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Specifications..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      Begin VB.CommandButton cmdaction 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Print Structures To .Txt File"
         Height          =   495
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   3015
      End
      Begin VB.OptionButton deselectall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "Deselect All"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   9
         Top             =   3000
         Width           =   1335
      End
      Begin VB.OptionButton selectall 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ListBox lsttable 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   1590
         ItemData        =   "Form1.frx":0442
         Left            =   240
         List            =   "Form1.frx":0444
         MultiSelect     =   1  'Simple
         TabIndex        =   5
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CommandButton cmddatabase 
         Caption         =   "..."
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   0
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Table Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   390
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   430
         Width           =   1365
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim Pk As New ADODB.Recordset
Dim Fk As New ADODB.Recordset
Dim table As New ADOX.table
Dim rs As New ADODB.Recordset
Dim ob As New ADOX.Catalog
Dim Table_Name$
Public Jetpassword, DatabasePath, provider
Dim Keys, Num, FileName, File, str, i, arr, DataType, PrintType, shel, loops, Flag, FilPath, TxtFile
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_Shownormal = 3

'Private Sub cboprovider_Click()
'provider = "Microsoft.Jet.Oledb." & Me.cboprovider.List(Me.cboprovider.ListIndex)
'If txtdata.Text <> "" Then OpenDatabase txtdata.Text
'End Sub

Private Sub cmdaction_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
  PrintType = "word"
  With cd
    .DialogTitle = "Select The Text File"
    .Filter = "Word Files (*.Doc)|*.doc"
    .FileName = ""
    .ShowOpen
    TxtFile = .FileName
    File = .FileName
  End With
Case 1
  PrintType = ""
  With cd
    .DialogTitle = "Select The Text File"
    .Filter = "Text Files (*.txt)|*.txt"
    .FileName = ""
    .ShowOpen
    TxtFile = .FileName
    File = .FileName
  End With
End Select
If Len(File) > 0 Then cmdgenerate_Click
End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmddatabase_Click()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  With cd
    .DialogTitle = "Select The Access Database Name"
    .Filter = "Access Files (*.mdb)|*.mdb"
    .ShowOpen
    txtdata.Text = .FileName
    FilPath = .FileName
    FileName = Mid(.FileTitle, 1, InStr(1, .FileTitle, ".") - 1)
  End With
  If cn.State = 1 Then
    cn.Close
  End If
  Screen.MousePointer = vbArrow
  DatabasePath = cd.FileName
  Me.lsttable.Clear
  TxtFile = ""
  PrintType = ""
  loops = 0
  If txtdata.Text <> "" Then OpenDatabase txtdata.Text
End Sub

Private Sub cmdgenerate_Click()
Dim x$
On Error GoTo jump
'check for select the table from listbox
Num = 0
For i = 0 To Me.lsttable.ListCount - 1
 If Me.lsttable.Selected(i) = True Then
    Num = Num + 1
 End If
Next
If Num < 1 Then
  MsgBox "Select The Table To Print The Structure", vbInformation
  Exit Sub
End If
'check the valid filename
If TxtFile = "" Then
  MsgBox "Select The Path To Save A File", vbInformation
  Exit Sub
ElseIf TxtFile <> "" Then
    
  If Right(File, 4) <> ".txt" And Right(File, 4) <> ".doc" Then
    MsgBox "Only .Txt And .Doc Extensions Files Are Allowed", vbInformation
    Exit Sub
  End If
End If

If Dir(TxtFile) <> "" Then
 If MsgBox(File & " Already Exists. Overwrite It ?", vbYesNo + vbQuestion) = vbYes Then
     Kill TxtFile
 Else
     Exit Sub
 End If
End If

Screen.MousePointer = vbHourglass
mhandle = FreeFile
Open TxtFile For Output As mhandle
  'print heading
  'Print #mhandle, Space(40) & UCase(FileName)
  'Print #mhandle, Space(40) & String(Len(FileName), "_")
  Print #mhandle, " Printing From " + FilPath
  Print #mhandle, ""
  Print #mhandle, "Printing Table Structures:"
  'print tables names only
   Dim Incre%
   For Num = 0 To Me.lsttable.ListCount - 1
     If Me.lsttable.Selected(Num) = True Then
       Incre = Incre + 1
       Print #mhandle, Incre & ") " & lsttable.List(Num)
       'Print #mhandle, ""
     End If
   Next
   Print #mhandle, ""
   Print #mhandle, ""
  'print table structures
  For Num = 0 To lsttable.ListCount - 1
     If lsttable.Selected(Num) = True Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from [" & lsttable.List(Num) & "]", cn, adOpenStatic, adLockOptimistic
        fldcount = rs.Fields.Count
        Print #mhandle, "Table Name : " & lsttable.List(Num) & Space(4) & "Total Fields : " & fldcount
        Print #mhandle, ""
        'Print #mhandle, String(90, "-")
        Print #mhandle, "Field Name" & Space(16) & "" & "DataType" & Space(5) & "" & "Description"
        Print #mhandle, "__________" & Space(16) & "" & "________" & Space(5) & "" & "___________"
        For i = 0 To fldcount - 1
            DoEvents
            size
            fldvalue = rs.Fields(i).Name
            If Len(fldvalue) < str Then
                'specify that datatype is text or not if it is then specify the width
                DataType = cType(rs.Fields(i).Type)
                If DataType = "Text" Then
                   DataType = cType(rs.Fields(i).Type) & "[" & rs.Fields(i).DefinedSize & "]"
                'Else
                   'DataType = cType(rs.Fields(i).Type)
                End If
                'CHECK PRIMARY KEY
                'Set Pk = cn.OpenSchema(adSchemaPrimaryKeys)
                'While Not Pk.EOF
                '   If lsttable.List(Num) = Pk.Fields("TABLE_NAME") Then
               '       If fldvalue = Pk.Fields("COLUMN_NAME") Then
               '          Keys = "PrimaryKey"
               '       End If
               '    End If
               ' Pk.MoveNext
               ' Wend
                'get field descrription
                x$ = GetFieldDesc_ADO(lsttable.List(Num), rs.Fields(i).Name)
               'CHECK FORIEGN KEY
                'Set Fk = cn.OpenSchema(adSchemaForeignKeys)
                'While Not Fk.EOF
                '  If lsttable.List(Num) = Fk.Fields("FK_TABLE_NAME") Then
                '    If fldvalue = Fk.Fields("FK_COLUMN_NAME") Then
                '       Keys = "ForeignKey" & "(" & Fk.Fields("PK_TABLE_NAME") & ")"
                '    End If
                '  End If
                'Fk.MoveNext
                'Wend
                'print structure
                Print #mhandle, "" & Space(1) & fldvalue & Space$(str - Len(fldvalue)) & "" & Space(1) & DataType & Space$(12 - Len(DataType)) & "" & Space(1) & x$ '& Space$((str - Len(Keys)) + 8) & ""
                Keys = ""
            Else
                Print #mhandle, fldvalue & "" & rs.Fields(i).Type
            End If
            'Print #mhandle, String(90, "-")
        Next i
        Print #mhandle, ""                ' Change line
        Print #mhandle, ""                ' Change line
     End If
  Next
  Print #mhandle, ""                ' Change line
  Close mhandle
  Screen.MousePointer = vbDefault
  'MsgBox "Structure Exported Sucessfully In " & vbCrLf _
  '       + "File : " & File, vbInformation
  'If PrintType = "word" Then
  '  MsgBox "Please Adjust The Left And Right " & vbCrLf _
  '         + "      Page Margins In The File", vbInformation
  '  PrintType = ""
  'End If
  shel = ShellExecute(O&, vbNullString, File, vbNullString, vbNullString, SW_Shownormal)
  Exit Sub
jump:
'  If Err.Number = 3251 Then
'    MsgBox "Your operating System doesn't support the [ " & provider & " ] provider" & vbCrLf _
'           + "Select another provider", vbInformation
'  Else
    MsgBox Err.Description, vbInformation
'  End If
  Screen.MousePointer = vbDefault
  Close mhandle
  Pk.CancelUpdate
  Fk.CancelUpdate
  rs.CancelUpdate
End Sub
Function GetFieldDesc_ADO(ByVal MyTableName As String, _
  ByVal MyFieldName As String)
   
   Dim MyDB As New ADOX.Catalog
   Dim MyTable As ADOX.table
   Dim MyField As ADOX.Column

   On Error GoTo Err_GetFieldDescription
   Me.Caption = MyTableName + ":" + MyFieldName
   MyDB.ActiveConnection = cn
   Set MyTable = MyDB.Tables(MyTableName)
   GetFieldDesc_ADO = MyTable.Columns(MyFieldName).Properties("Description")
   
   Set MyDB = Nothing

Bye_GetFieldDescription:
   Exit Function

Err_GetFieldDescription:
   Beep
   MsgBox Err.Description, vbExclamation
   GetFieldDescription = Null
   Resume Bye_GetFieldDescription

End Function
Sub size()
 str = 25
End Sub

Public Function cType(ByVal Value As ADOX.DataTypeEnum) As String
  Select Case Value
    Case adTinyInt: cType = "TinyInt"
    Case adSmallInt: cType = "SmallInt"
    Case adInteger: cType = "Number"
    Case adBigInt: cType = "BigInt"
    Case adUnsignedTinyInt: cType = "UnsignedTinyInt"
    Case adUnsignedSmallInt: cType = "UnsignedSmallInt"
    Case adUnsignedInt: cType = "UnsignedInt"
    Case adUnsignedBigInt: cType = "UnsignedBigInt"
    Case adSingle: cType = "Single"
    Case adDouble: cType = "Double"
    Case adCurrency: cType = "Currency"
    Case adDecimal: cType = "Decimal"
    Case adNumeric: cType = "Numeric"
    Case adBoolean: cType = "Boolean"
    Case adUserDefined: cType = "UserDefined"
    Case adVariant: cType = "Variant"
    Case adGUID: cType = "GUID"
    Case adDate: cType = "Date/Time"
    Case adDBDate: cType = "Date/Time"
    Case adDBTime: cType = "Date/Time"
    Case adDBTimeStamp: cType = "Date/Time"
    Case adBSTR: cType = "BSTR"
    Case adChar: cType = "Text"
    Case adVarChar: cType = "Text"
    Case adLongVarChar: cType = "Text"
    Case adWChar: cType = "Text"
    Case adVarWChar: cType = "Text"
    Case adLongVarWChar: cType = "Memo"
    Case adBinary: cType = "adBinary"
    Case adVarBinary: cType = "adVarBinary"
    Case adLongVarBinary: cType = "OLE Object"
    Case Else: cType = Value
  End Select
End Function

Private Sub deselectall_Click()
For i = 0 To Me.lsttable.ListCount - 1
 Me.lsttable.Selected(i) = False
Next
End Sub

Private Sub Form_Load()
'MsgBox LoadResString(1), vbExclamation
provider = "Microsoft.Jet.Oledb.4.0"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
'frmAbout.Show
End Sub

Private Sub selectall_Click()
For i = 0 To Me.lsttable.ListCount - 1
 Me.lsttable.Selected(i) = True
Next
End Sub

Private Sub Timer1_Timer()
If Caption = "" Then
  Caption = " PRINT STRUCTURE"
ElseIf Caption = " PRINT STRUCTURE" Then
  Caption = ""
Else
  Caption = ""
End If
End Sub

Public Sub OpenDatabase(Files As String)
On Error Resume Next
Do
  loops = loops + 1
  If cn.State = 1 Then cn.Close
  cn.provider = provider
  cn.CursorLocation = adUseClient
  cn.Properties("Data Source") = Files
  cn.Properties("Jet OLEDB:Database Password") = Jetpassword
  cn.Open
  If Err.Number = 0 Then
    loops = 0
  ElseIf Err.Number = -2147217843 And loops = 1 Then
    Jetpassword = GetAccess97Password(Files) 'Password 97 auto login
  ElseIf Err.Number = -2147217843 And loops > 1 Then
    Jetpassword = ""
    'Unload Table
    Password.Show 'password 2000 'prompt dialog
    If loops > 2 Then
      'loops = 0
      Password.Hide
      MsgBox "Not A Valid Password", vbExclamation
      Password.Show
      Password.txtPassword.Text = ""
      Password.txtPassword.SetFocus
    End If
    Exit Sub
  Else
    MsgBox "Unable To Open Database...", vbCritical
    loops = 0
    Exit Sub
  End If
Loop While (loops > 0)
 Unload Password
 pass = Jetpassword
 loops = 0
 Jetpassword = ""
  ob.ActiveConnection = cn
  Me.lsttable.Clear
  Screen.MousePointer = vbHourglass
  For Each table In ob.Tables
    If table.Type = "TABLE" Then
      If UCase(Left(table.Name, 4)) <> "MSYS" Then
         lsttable.AddItem table.Name
      End If
    End If
  Next
  Screen.MousePointer = vbArrow
End Sub

Public Function GetAccess97Password(ByVal FileName As String) As String
On Error GoTo errHandler
Dim ch(18) As Byte
Dim x As Integer
Dim Sec
  GetAccess97Password = ""
  If Trim(FileName) = "" Then Exit Function
' Used integers instead of hex :-)  Easier to read
  Sec = Array(0, 134, 251, 236, 55, 93, 68, 156, 250, 198, 94, 40, 230, 19, 182, 138, 96, 84)
  
  Open FileName For Binary Access Read As #1 Len = 18
  Get #1, &H42, ch
  Close #1
  
  For x = 1 To 17
    GetAccess97Password = GetAccess97Password & Chr$(ch(x) Xor Sec(x))
  Next x
  GetAccess97Password = Replace(GetAccess97Password, Chr$(0), "")
Exit Function
errHandler:
  MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
  Exit Function
  Resume
End Function

