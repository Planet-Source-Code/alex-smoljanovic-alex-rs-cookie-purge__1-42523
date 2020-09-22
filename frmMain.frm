VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cookie Purge"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Invert Selection"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select None"
      Height          =   375
      Left            =   1140
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select All"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete Cookies"
      Height          =   375
      Left            =   4020
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin ComctlLib.ListView lstCookies 
      Height          =   2235
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3942
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cookie Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cookie Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtCookieDir 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   4035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cookies:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cookie Directory:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'This project was explicitly developed for
'PSC(Planet Source Code) Users as an Open Source Project.
'This code is the property of it's author.
'
'If you compile this application you may not redistribute it.
'However, you may use any of this code in you're own application(s).
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************

Dim CookieDir$ 'dimensionalize variable as string type

Private Sub Command1_Click()
On Error Resume Next 'on the event of an error, resume execution on the next line of this procedure
Dim i% 'dimensionalize i as integer type
 If MsgBox("Are you sure you wish to delete the selected cookies?", vbQuestion + vbYesNo, "Delete Cookies") = vbYes Then
 'request user confirmation
  For i = lstCookies.ListItems.Count To 1 Step -1
  'for next loop, i evaluates to the number of Items in the ListItem collection
  'loops until i = 1 decrementing i by one each iteration
   If lstCookies.ListItems(i).Text <> "" And lstCookies.ListItems(i).Selected = True Then
   'if the current items text property is unequal to "", and the item is selected then...
    Err.Clear 'call object err's Clear method to destroy the current error infor, if any
     Kill lstCookies.ListItems(i).Key
     'purge the cookie from the hard drive, note: the cookies path is stored in the list items key property
      If Err.Number = 75 Then 'if a File/Path access error occured, attempt to compensate...
       SetAttr lstCookies.ListItems(i).Key, vbNormal
       'remove any extended file attributes...
        Err.Clear 'clear the current error information
         Kill lstCookies.ListItems(i).Key 'attempt to purge the file once more
          If Err.Number = 0 Then lstCookies.ListItems.Remove i
          'if no error occured, remove the list item representing the cookie from the listview control
      ElseIf Err.Number = 0 Then lstCookies.ListItems.Remove i
      'if no error occured, remove the list item representing the cookie from the listview control
      End If
   End If
  Next i 'next loop iteration; increment i; check loop conditions
 End If
End Sub

Private Sub Command2_Click()
Dim i% 'dimensionalize i as integer data type
 For i = 1 To lstCookies.ListItems.Count
 'for next loop, loop through each list item in the listitems collection
  lstCookies.ListItems(i).Selected = True
  'update the list item's selected property to true
 Next i 'next iteration
End Sub

Private Sub Command3_Click()
Dim i% 'dimensionalize i as integer data type
 For i = 1 To lstCookies.ListItems.Count
 'for next loop, loop through each list item in the listitems collection
  lstCookies.ListItems(i).Selected = False
  'update the list item's selected property to false
 Next i 'next iteration
End Sub

Private Sub Command4_Click()
Dim i% 'dimensionalize i as integer data type
 For i = 1 To lstCookies.ListItems.Count
 'for next loop, loop through each list item in the listitems collection
  lstCookies.ListItems(i).Selected = Not (lstCookies.ListItems(i).Selected)
  'perform a logical not operation on the value of the list items selected property to invert its property value(False evaluates to true, true evaluates to false)
 Next i 'next iteration
End Sub

Private Sub Form_Load()
 CookieDir$ = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cookies")
 'See GetString function for more info...
  txtCookieDir = CookieDir 'update the text boxes Text property with the Cookie Directory
   EnumCookies CookieDir 'see EnumCookies for more info...
End Sub

Public Sub EnumCookies(ByVal Path$)
Dim buffer$, itmX As ListItem
'dimensionalize buffer as string data type, itmX as ListItem structure
 Path = Path & IIf(Right$(Path, 1) = "\", "", "\")
 'Use IIf operator to conditionally append the character "\" to the Path variable if the last character of Path isn't "\"
  buffer = Dir(Path, vbDirectory) 'Change Directory to the Cookie Directory,
  'additionally return the first file in the specified path
   If buffer <> "" Then 'if the directory contains files then...
    Do While buffer <> ""
    'Do While loop; loop until variable buffer evaluates to ""
    DoEvents 'yield execution to other asynchronously processing procedures...
     buffer = Dir(, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
     'return the next file(of any file attribute) in the previously established directory
      If buffer <> "" Then 'if buffer is unequal to ""(next file was returned)
       Set itmX = lstCookies.ListItems.Add(, Path & buffer, buffer)
       'initialize itmX with the ListItem type return of the Add method
        itmX.SubItems(1) = Path & buffer
        'update the items subitem text
      End If
    Loop 'check loop conditions; loop if conditions evaluate to true(buffer unequal to "")
     On Error Resume Next 'on the event of an error resume next
     'this will assure that even if there were no file in the cookies directory an Object variable or with block not set error will not occur
      lstCookies.ListItems(lstCookies.ListItems.Count).Selected = True
      'update the last listitem item in the listitems collection's Selected property to true(selected)
   End If
End Sub
