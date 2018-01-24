VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Code As String
Private Offset As Integer

Public Sub Initialize(CodeToParse As String)
    Offset = 0
    Code = CodeToParse
End Sub

Private Sub GoNextChar()
    Offset = Offset + 1
End Sub

Private Function IsEOF()
    IsEOF = Offset >= Len(Code)
End Function

Private Function GetCurrentChar()
    GetCurrentChar = Mid(Code, Offset, 1)
End Function

Private Function ParseSymbol() As Symbol
    Dim Keyword As String
    
    Do While Not IsEOF()
        Keyword = Keyword + GetCurrentChar()
        Call GoNextChar
    Loop
    
    ParseSymbol = CreateSymbol(Keyword)
End Function

Private Function ParseProgram() As Variant()
    Dim Program() As Variant
        
    Do While Not IsEOF()
        Call SkipWhiteSpaces
        ReDim Preserve Program(0 To UBound(Program) + 1) As Variant
        Program(UBound(Program)) = ParseExpression
        Call SkipWhiteSpaces
    Loop
    
    ParseProgram = Program
End Function

Private Sub SkipWhiteSpaces()

End Sub

Private Function ParseExpression() As Variant

End Function

Public Function Parse()
    Parse = Code
End Function
