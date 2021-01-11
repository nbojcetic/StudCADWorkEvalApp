Public Class Part
    Public Property path As Attribute
    Public Property name As Attribute
    Public Property type As Attribute
    Public Property units As Attribute
    Public Property featureNo As Attribute
    Public Property material As Attribute
    Public Property mass As Attribute
    Public Property density As Attribute
    Public Property volume As Attribute
    Public Property MaxPoints As String
    Public Property envelope As List(Of Attribute)
    Public Property features As List(Of Feature)
    Public Property equations As List(Of Equation)
    Public Property properties As List(Of Attribute)

    Public Sub New()
        path = New Attribute()
        name = New Attribute()
        type = New Attribute()
        units = New Attribute()
        featureNo = New Attribute()
        material = New Attribute()
        mass = New Attribute()
        density = New Attribute()
        volume = New Attribute()
        MaxPoints = 0
    End Sub
End Class