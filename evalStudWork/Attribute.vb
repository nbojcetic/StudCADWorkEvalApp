Public Class Attribute
    Public Property name As String
    Public Property value As String
    Public Property rulepoints As String
    Public Property ruletype As String
    Public Property rulecheck As String
    Public Property rulevalue As String
    Public Property ruledescription As String

    Public Sub New()
        name = ""
        value = ""
        ruletype = ""
        rulecheck = ""
        rulevalue = ""
        ruledescription = ""
        rulepoints = ""
    End Sub

    Public Sub New(ByVal inName As String, ByVal inValue As String, inRPoints As String, inRChk As String,
                   inRType As String, inRValue As String, inRDesc As String)
        name = inName
        value = inValue
        rulepoints = inRPoints
        rulecheck = inRChk
        ruletype = inRType
        rulevalue = inRValue
        ruledescription = inRDesc
    End Sub
End Class