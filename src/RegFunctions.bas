Attribute VB_Name = "RegFunctions"
Option Explicit
Private m_lngRetVal As Long
Private Const REG_NONE As Long = 0
Private Const REG_SZ As Long = 1
Private Const REG_EXPAND_SZ As Long = 2
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4
Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Private Const REG_DWORD_BIG_ENDIAN As Long = 5
Private Const REG_LINK As Long = 6
Private Const REG_MULTI_SZ As Long = 7
Private Const REG_RESOURCE_LIST As Long = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_ALL_ACCESS As Long = &H3F
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_NO_MORE_ITEMS As Long = 259
Private Const REG_OPTION_NON_VOLATILE As Long = 0
Private Const REG_OPTION_VOLATILE As Long = &H1
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Function regQuery_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String) As Variant
  Dim intPosition As Integer
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngBufferSize As Long
  Dim lngBuffer As Long
  Dim strBuffer As String
  lngKeyHandle = 0
  lngBufferSize = 0
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)
      Exit Function
  End If
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal 0&, lngBufferSize)
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)
      Exit Function
  End If
  Select Case lngDataType
         Case REG_SZ:
              strBuffer = Space(lngBufferSize)
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, ByVal strBuffer, lngBufferSize)
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  intPosition = InStr(1, strBuffer, Chr(0))
                  If intPosition > 0 Then
                      regQuery_A_Key = Left(strBuffer, intPosition - 1)
                  Else
                      regQuery_A_Key = strBuffer
                  End If
              End If
         Case REG_DWORD:
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  regQuery_A_Key = lngBuffer
              End If
         Case Else:
              regQuery_A_Key = ""
  End Select
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function
