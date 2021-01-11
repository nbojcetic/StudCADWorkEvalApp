Public Class TableData
    Private name As String
    Private value As String
    Private rpoints As String
    Private rcheck As String
    Private rtype As String
    Private rvalue As String
    Private rdescr As String
    Private section As String

    Public Property nameProp() As String
        Get
            Return name
        End Get
        Set(ByVal inVal As String)
            name = inVal
        End Set
    End Property

    Public Property valueProp() As String
        Get
            Return value
        End Get
        Set(ByVal inVal As String)
            value = inVal
        End Set
    End Property

    Public Property rpointsProp() As String
        Get
            Return rpoints
        End Get
        Set(ByVal inVal As String)
            rpoints = inVal
        End Set
    End Property

    Public Property rcheckProp() As String
        Get
            Return rcheck
        End Get
        Set(ByVal inVal As String)
            rcheck = inVal
        End Set
    End Property

    Public Property rtypeProp() As String
        Get
            Return rtype
        End Get
        Set(ByVal inVal As String)
            rtype = inVal
        End Set
    End Property

    Public Property rvalueProp() As String
        Get
            Return rvalue
        End Get
        Set(ByVal inVal As String)
            rvalue = inVal
        End Set
    End Property

    Public Property rdescrProp() As String
        Get
            Return rdescr
        End Get
        Set(ByVal inVal As String)
            rdescr = inVal
        End Set
    End Property

    Public Property sectionProp() As String
        Get
            Return section
        End Get
        Set(ByVal inVal As String)
            section = inVal
        End Set
    End Property

    Public Sub New(inName As String, inValue As String, inRpoints As String, inRcheck As String,
                         inRtype As String, inRvalue As String, inRdescr As String, inSection As String)
        name = inName
        value = inValue
        rpoints = inRpoints
        rcheck = inRcheck
        rtype = inRtype
        rvalue = inRvalue
        rdescr = inRdescr
        section = inSection
    End Sub
End Class
