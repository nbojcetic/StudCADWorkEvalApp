Public Class Equation
    Public Property name As Attribute
    Public Property value As Attribute
    Public Property status As Attribute
    Public Property isglobal As Attribute

    Public Sub New()
        name = New Attribute()
        value = New Attribute()
        status = New Attribute()
        isglobal = New Attribute()
    End Sub
End Class