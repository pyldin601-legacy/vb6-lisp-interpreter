Attribute VB_Name = "Types"
Type Symbol
    Name As String
End Type

Public Function CreateSymbol(Name As String) As Symbol
    Dim SymbolInstance As Symbol
    
    SymbolInstance.Name = Name
    
    CreateSymbol = SymbolInstance
End Function
