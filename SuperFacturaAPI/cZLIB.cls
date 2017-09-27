VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZLIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
'
' Fuente: https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64920&lngWId=1
'
'--------------------------------------------------------------------------------
'    Component  : cZLIB
'    Project    : axHTTP
'    Author     : James Johnston, TechKnow Professional Services
'                 jjohnston@techknowpro.com, http://www.techknowpro.com
'
'    Description: Allows for easy compress/uncompress of string data using the
'                 Deflate Algorithm.  You can easily wrap the data with a
'                 ZLib or GZip header by specifying the zMode when calling the
'                 functions.  vbzlib1.dll is a modified version of zlib1.dll.
'                 http://www.zlib.org
'
'    Modified   : 12/21/2005
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Enum zMode
    Z_DEFLATE = -15     'Will compress/uncompress without zlib or gzip wrapper
    Z_ZLIB = 15         'Will compress/uncompress data with zlib wrapper
    Z_GZIP = 15 + 16    'Will compress/uncompress data with gzip wrapper
    Z_AUTO = 15 + 32    'Will compress data with zlib wrapper but uncompress data with zlib or gzip wrapper
End Enum

Public Enum zCode
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

Private Declare Function adler32 Lib "vbzlib1.dll" (ByVal adler As Long, ByRef buf As Any, ByVal length As Long) As Long
Private Declare Function crc32 Lib "vbzlib1.dll" (ByVal crc As Long, ByRef buf As Any, ByVal length As Long) As Long
Private Declare Function compress Lib "vbzlib1.dll" (ByRef dest As Any, ByRef destLen As Long, ByRef Source As Any, ByVal sourceLen As Long, ByVal windowBits As Long) As Long
Private Declare Function compressBound Lib "vbzlib1.dll" (ByVal sourceLen As Long) As Long
Private Declare Function uncompress Lib "vbzlib1.dll" (ByRef dest As Any, ByRef destLen As Long, ByRef Source As Any, ByVal sourceLen As Long, ByVal windowBits As Long) As Long

Public Function GetAdler32(ByVal Data As String) As Long
    Dim crc As Long
    
    ' Get initial value
    crc = adler32(0, ByVal 0&, 0)
    crc = adler32(crc, Data, Len(Data))
    
    GetAdler32 = crc
End Function

Public Function GetCRC32(ByVal Data As String) As Long
    Dim crc As Long
    
    ' Get initial value
    crc = crc32(0, ByVal 0&, 0)
    crc = crc32(crc, Data, Len(Data))
    
    GetCRC32 = crc
End Function

Public Function CompressString(ByRef Data As String, ByVal Mode As zMode) As Long
    Dim sInput() As Byte, lInput As Long
    Dim sOutput() As Byte, lOutput As Long
    Dim RetVal As zCode
    
    On Error GoTo ErrorHandler
    
    'Convert Data to byte array
    sInput = StrConv(Data, vbFromUnicode)
    lInput = UBound(sInput) + 1
    If (lInput > 0) Then
        lOutput = compressBound(lInput)
        ReDim sOutput(0 To (lOutput - 1)) As Byte
        
        RetVal = compress(sOutput(0), lOutput, sInput(0), lInput, Mode)
        
        'Was compress successful?
        If (RetVal = Z_OK) Then
            Data = Left$(StrConv(sOutput, vbUnicode), lOutput)
            CompressString = RetVal
        Else
            CompressString = Z_ERRNO
        End If
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error cZLIB.CompressString: " & err.Number, err.Description
    CompressString = err.Number
End Function


Public Function UncompressString(ByRef Data As String, ByVal Mode As zMode) As Long
    Dim sInput() As Byte, lInput As Long, sDeflate As String
    Dim sOutput() As Byte, lOutput As Long
    Dim RetVal As zCode
    'Dim crc As Long
    
    Dim gzID1 As Byte, gzID2 As Byte, gzCM As Byte
    Dim gzFLG As Byte, gzXFL As Byte, gzOS As Byte
    Dim gzCRC32 As Long, gzISIZE As Long
    Dim gzFTEXT As Boolean, gzFHCRC As Boolean, gzFEXTRA As Boolean
    Dim gzFNAME As Boolean, gzFCOMMENT As Boolean
    Dim Offset As Long
    
    On Error GoTo ErrorHandler
    
    sInput = StrConv(Data, vbFromUnicode)
    lInput = UBound(sInput) + 1
    
    If lInput = 0 Then Exit Function
    
    If Mode = Z_GZIP Then
        'Read Header
        gzID1 = sInput(0) '31
        gzID2 = sInput(1) '139
        gzCM = sInput(2)  '8 = deflate
        gzFLG = sInput(3) '0
        gzXFL = sInput(8) '2 = max, 4 = fast
        gzOS = sInput(9)  '255 = unknown
        
        'Decode Flags
        gzFTEXT = gzFLG And 1
        gzFHCRC = gzFLG And 2
        gzFEXTRA = gzFLG And 4
        gzFNAME = gzFLG And 8
        gzFCOMMENT = gzFLG And 16
        
        Offset = lInput - 8
        gzCRC32 = CLng(sInput(Offset)) + (CLng(sInput(Offset + 1)) * 256)
        
        Offset = lInput - 4
        gzISIZE = CLng(sInput(Offset)) + (CLng(sInput(Offset + 1)) * 256)
    Else
        gzISIZE = lInput * 3
    End If
    
    If lInput > 0 And gzISIZE > 0 Then
        
        'Make sure the output buffer is big enough to hold the uncompressed data
        Do
            lOutput = gzISIZE
            ReDim sOutput(0 To (lOutput - 1)) As Byte
            RetVal = uncompress(sOutput(0), lOutput, sInput(0), lInput, Mode)
            gzISIZE = gzISIZE + lInput
        Loop While RetVal = Z_BUF_ERROR
        
        'Was Uncompress Successful?
        If (RetVal = Z_OK) Then
            Data = Left$(StrConv(sOutput, vbUnicode), lOutput)
            UncompressString = RetVal
        Else
            UncompressString = Z_ERRNO
        End If
    End If
    Exit Function

ErrorHandler:
    Debug.Print "Error cZLIB.UncompressString: " & err.Number, err.Description
    UncompressString = err.Number
End Function


