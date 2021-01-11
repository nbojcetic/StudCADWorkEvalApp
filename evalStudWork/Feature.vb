Public Class Feature
    Public Property name As Attribute
    Public Property type As Attribute
    Public Property base As Attribute
    Public Property frozen As Attribute
    Public Property suppressed As Attribute
    Public Property attribs As List(Of Attribute)
    Public Property dimensions As List(Of Dimension)
    Public Property sketches As List(Of Sketch)

    Public Sub New()
        name = New Attribute()
        type = New Attribute()
        base = New Attribute()
        frozen = New Attribute()
        suppressed = New Attribute()
    End Sub
End Class
