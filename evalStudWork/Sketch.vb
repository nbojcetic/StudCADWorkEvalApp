Public Class Sketch
    Public Property name As Attribute
    Public Property type As Attribute
    Public Property status As Attribute
    Public Property constrained As Attribute
    Public Property geometries As List(Of Geometry)

    Public Sub New()
        name = New Attribute()
        type = New Attribute()
        status = New Attribute()
        constrained = New Attribute()
    End Sub
End Class
