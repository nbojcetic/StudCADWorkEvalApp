Public Class Dimension
    Public Property name As Attribute
    Public Property type As Attribute
    Public Property value As Attribute

    Public Sub New()
        name = New Attribute()
        value = New Attribute()
        type = New Attribute()
    End Sub

    Public Sub New(ByVal inName As String, ByVal inValue As String, inRPoints As String, inRChk As String,
                   inRType As String, inRValue As String, inRDesc As String)
        name = New Attribute("Dimension Name", inName, inRPoints, inRChk, inRType, inRValue, inRDesc)
        value = New Attribute("Dimension Value", inValue, inRPoints, inRChk, inRType, inRValue, inRDesc)
        type = New Attribute("Dimension Type", inRType, inRPoints, inRChk, inRType, inRValue, inRDesc)
    End Sub
End Class
