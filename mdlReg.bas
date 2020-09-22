Attribute VB_Name = "modReg"
Option Explicit
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
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
 Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
 Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
 Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
 Private Const FORMAT_MESSAGE_FROM_STRING = &H400
 Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
 Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
 Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_EXPAND_SZ = 2

Const ERROR_SUCCESS = 0&

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function RegQueryStringValue(ByVal hKey&, ByVal strKeyName$) As String
Dim rLng&, rKeyType, buffer$, rKeyLength&, lBuffer&
'dimensionalize rLng& as long type, rKeyType as long type, buffer$ as string type, lDataBuSize as long type, lBuffer as integer type
rLng& = RegQueryValueEx(hKey, strKeyName, 0, rKeyType, ByVal 0, rKeyLength)
'function retrieves the type and data for a specified value name associated with an open registry key
 If rLng& = ERROR_SUCCESS Then
 'if function was successful rLng& evaluates to ERROR_SUCCESS
  If rKeyType = REG_SZ Or rKeyType = REG_EXPAND_SZ Then
  'if rKeyType evaluates to REG_SZ(String) then...
   buffer$ = String(rKeyLength, Chr$(0))
   'allocate memory to buffer$ variable
    rLng& = RegQueryValueEx(hKey, strKeyName, 0, 0, ByVal buffer$, rKeyLength)
    'initialize buffer$ with the key's data
     If rLng& = ERROR_SUCCESS Then
     'if the function was successful then...
      RegQueryStringValue = Left$(buffer$, InStr(1, buffer$, Chr$(0)) - 1)
      'remove the nullterminating characters...(returns only the length of the key)
     End If
  ElseIf rKeyType = REG_BINARY Then
  'if rKeyType(the specified key's key type) evaluates to REG_BINARY(DWORD) then...
    rLng& = RegQueryValueEx(hKey, strKeyName, 0, 0, lBuffer, rKeyLength)
    'initialize lBuffer with the key's data
     If rLng& = ERROR_SUCCESS Then
     'if the function was successful then...
      RegQueryStringValue = CStr(lBuffer)
      'return the keys data
     End If
  End If
 Else
  MsgBox FormatSystemError
 End If
End Function

Function GetString(hKey&, strPath$, strValue$)
Dim rRes& 'dimensionalize rRes as long type
 RegOpenKey hKey, strPath, rRes
 'opens a handle to the key
  GetString = RegQueryStringValue(rRes, strValue)
  'see RegQueryStringValue for more info...
   RegCloseKey rRes 'close the keys handle
End Function

Sub SaveString(hKey&, strPath$, strValue$, strData$)
Dim rRes& 'dimensionalize rRes as long type
 RegCreateKey hKey, strPath, rRes
 'function creates the specified registry key, if the key already exists the function opens it
  RegSetValueEx rRes, strValue, 0, REG_SZ, ByVal strData, Len(strData)
  'set the data and the type(REG_SZ[String]) of the specified key
   RegCloseKey rRes 'close the handle to the key
End Sub

Sub SaveStringLong(hKey&, strPath$, strValue$, strData$)
Dim rRes& 'dimensionalize rRes as long type
 RegCreateKey hKey, strPath, rRes
 'function creates the specified registry key, if the key already exists the function opens it
  RegSetValueEx rRes, strValue, 0, REG_BINARY, CByte(strData), 4
  'set the data and the type(REG_BINARY[Binary(DWORD)]) of the specified key
   RegCloseKey rRes 'close the keys handle
End Sub

Sub DelSetting(hKey&, strPath$, strValue$)
Dim rRes& 'dimensionalize rRes as long type
 RegCreateKey hKey, strPath, rRes
 'function creates the specified registry key, if the key already exists the *function opens it*
  RegDeleteValue rRes, strValue
  'deletes the specified registry key from the registry
   RegCloseKey rRes 'close the keys handle
End Sub

'*********************************************
'Used while debugging....
Private Function FormatSystemError() As String
 FormatSystemError = Space$(200)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, 0&, FormatSystemError, 200, ByVal 0&
   FormatSystemError = RTrim$(FormatSystemError)
End Function
'*********************************************
