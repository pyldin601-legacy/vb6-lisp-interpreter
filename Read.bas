Attribute VB_Name = "Read"
Function ParseCode(Code As String) As Variant()
    Dim ParserInstance As New Parser
    Call ParserInstance.Initialize(Code)
End Function

