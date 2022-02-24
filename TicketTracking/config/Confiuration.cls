VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Confiuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function getConfigFrom(ByVal fileContent As String, configName As String)
    
    Dim charIndex As Integer
    Dim returnVal As String
    
    On Error Resume Next
    
    charIndex = InStr(1, fileContent, configName, vbTextCompare)
    'returnVal = Mid$(fileContent, charIndex, Len(fileContent))
    returnVal = Mid$(fileContent, charIndex)
 
    charIndex = InStr(1, returnVal, vbCrLf, vbTextCompare)
    returnVal = Mid$(returnVal, 1, charIndex - 1)
    
    charIndex = InStr(1, returnVal, "=", vbTextCompare)
    'returnVal = Mid$(returnVal, charIndex + 1, Len(returnVal))
    returnVal = Mid$(returnVal, charIndex + 1)
    
    getConfigFrom = returnVal
    
End Function

