VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'[Declares]
  Declare Function compress _
    Lib "zlib32.dll" _
     (ByVal compr As String, _
      comprLen As Any, _
      ByVal buf As String, _
      ByVal buflen As Long) _
        As Integer
  Declare Function uncompress _
    Lib "zlib32.dll" _
     (ByVal uncompr As String, _
      uncomprLen As Any, _
      ByVal compr As String, _
      ByVal lcompr As Long) _
        As Long
  Declare Function gzopen _
    Lib "zlib32.dll" _
     (ByVal file As String, _
      ByVal mode As String) _
        As Long
  Declare Function gzread _
    Lib "zlib32.dll" _
     (ByVal file As Long, _
      ByVal uncompr As String, _
      ByVal uncomprLen As Long) _
        As Long
  Declare Function gzwrite _
    Lib "zlib32.dll" _
     (ByVal file As Long, _
      ByVal uncompr As String, _
      ByVal uncomprLen As Long) _
        As Long
  Declare Function gzclose _
    Lib "zlib32.dll" _
     (ByVal file As Long) _
        As Long

