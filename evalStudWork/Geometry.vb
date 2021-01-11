Public Class Geometry
    Public Property name As Attribute
    Public Property attribs As List(Of Attribute)
    Public Property relations As List(Of Attribute)

    Public Sub New()
        name = New Attribute()
    End Sub
End Class
