VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Read-only parameters that can (or not) be object references, although are references.
'Return same values as StrComp().
Public Function CompareValues(ByRef pvValue1 As Variant, ByRef pvValue2 As Variant) As Integer

End Function
