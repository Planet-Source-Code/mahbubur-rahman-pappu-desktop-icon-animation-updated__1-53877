VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WritePrivateProfileString _
    Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationname As String, ByVal lpKeyName As Any, _
            ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileInt Lib _
    "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, _
        ByVal lpKeyName As String, ByVal nDefault As Long, _
            ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
        ByVal lpKeyName As String, ByVal lpDefault As String, _
            ByVal lpReturnedString As String, ByVal nSize As Long, _
                ByVal lpFileName As String) As Long

Private m_strFile As String ' for file name
Private m_AppName As String '
Public Property Get FileName() As String
    FileName = m_strFile
End Property

Public Property Let FileName(ByVal strFile As String)
    m_strFile = Trim(strFile)
End Property

Private Sub Class_Initialize()
    m_strFile = ""
    m_AppName = ""
End Sub

Public Property Get ApplicationKey() As String
    ApplicationKey = m_AppName
End Property

Public Property Let ApplicationKey(ByVal strAppKey As String)
    m_AppName = Trim(strAppKey)
End Property
Private Function fConvert(str As String) As String
Dim i As Integer
Dim strOP As String
Dim intAsc As Integer

    str = Trim(str)
    
    strOP = ""
    
    For i = 1 To Len(str)
        intAsc = Asc(Mid(str, i, 1))
        If intAsc > 31 And intAsc < 127 Then _
            strOP = strOP & Chr(intAsc)
    Next
    fConvert = strOP
End Function

Public Function GetValue(ByVal keyName As String, ByVal strDefault As String) As String
Dim strVal As String * 100
Dim ret As Long
    ret = GetPrivateProfileString(m_AppName, keyName, strDefault, strVal, Len(strVal), m_strFile)
    If ret <> 0 Then
        GetValue = fConvert(strVal)
    Else
        GetValue = "Error"
    End If
End Function
Public Function SetValue(ByVal keyName As String, ByVal strValue As String) As Boolean
Dim ret As Long

ret = WritePrivateProfileString(m_AppName, keyName, strValue, m_strFile)
If ret = 0 Then
    SetValue = False
    Exit Function
End If
SetValue = True

End Function

