Public Class reportItem
    Private name As String
    Private orgVal As String
    Private masterVal As String
    Private points As String
    Private dscr As String
    Public Property nameProp() As String
        Get
            Return name
        End Get
        Set(ByVal inVal As String)
            name = inVal
        End Set
    End Property
    Public Property orgValProp() As String
        Get
            Return orgVal
        End Get
        Set(ByVal inVal As String)
            orgVal = inVal
        End Set
    End Property
    Public Property masterValProp() As String
        Get
            Return masterVal
        End Get
        Set(ByVal inVal As String)
            masterVal = inVal
        End Set
    End Property
    Public Property pointsProp() As String
        Get
            Return points
        End Get
        Set(ByVal inVal As String)
            points = inVal
        End Set
    End Property
    Public Property dscrProp() As String
        Get
            Return dscr
        End Get
        Set(ByVal inVal As String)
            dscr = inVal
        End Set
    End Property

    Public Sub New(inNameVal As String, inOrgVal As String, inMasterVal As String,
                   inPoints As String, inDscr As String)
        orgVal = inOrgVal
        masterVal = inMasterVal
        points = inPoints
        dscr = inDscr
    End Sub
End Class
