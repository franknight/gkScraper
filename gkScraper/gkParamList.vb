Public Class gkParamList

    Private p As SortedList

    Public Sub New()
        p = New SortedList
    End Sub

    Public Sub Add(ByVal Name, ByVal Value)
        p.Add(Name, Value)
    End Sub

    Default Public Property Item(ByVal Name As String) As Object
        Get
            Return p(Name)
        End Get
        Set(ByVal value As Object)
            Dim o As Object
            o = p(Name)
            If Not o Is Nothing Then
                p(Name) = value
            Else
                p.Add(Name, value)
            End If
        End Set
    End Property
    Public Property GetValue(ByVal Name As String)
        Get
            Return p(Name)
        End Get
        Set(ByVal value)
            p(Name) = value
        End Set
    End Property

    Public ReadOnly Property Keys() As String()
        Get
            Return p.Keys
        End Get
    End Property

    Public Function GetFieldValuesPair() As IDictionaryEnumerator
        Return p.GetEnumerator
    End Function

    Public Class ResultComparer
        Implements IComparer

        Public Enum SortOrder
            Asc
            Desc
        End Enum

        Private Field2Compare As String
        Private pSortOrder As SortOrder

        Public Sub New(ByVal FieldName As String, Optional ByVal order As gkParamList.ResultComparer.SortOrder = SortOrder.Asc)
            Field2Compare = FieldName
            pSortOrder = order
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            If Not TypeOf x Is gkParamList Then
                Throw New System.ArgumentException
            End If
            If Not TypeOf y Is gkParamList Then
                Throw New System.ArgumentException
            End If

            Dim xx As Object
            Dim yy As Object
            xx = x(Me.Field2Compare)
            yy = y(Me.Field2Compare)
            If pSortOrder = SortOrder.Asc Then
                Return xx.compareto(yy)
            Else
                Return yy.compareto(xx)
            End If

        End Function
    End Class

End Class
