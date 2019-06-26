Imports System.Data
Imports System.Data.OleDb

'"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=database.accdb;Persist Security Info=False;"

Public Class DBAccessHelper

    Private pc As OleDbConnection
    Private cmd As OleDbCommand

    Private pCmdType As DBCommandTypes
    Private pTable As String
    Private pFields As SortedList
    Private pValues As ArrayList
    Private pWhere As String

    Public IgnorePrimaryKeyViolation As Boolean

    <DebuggerStepThroughAttribute()> _
    Public Sub New()
        pc = New OleDbConnection
        pFields = New SortedList
        pValues = New ArrayList
    End Sub
    <DebuggerStepThroughAttribute()> _
    Public Sub New(ByVal dba As DBAccessHelper)
        pc = dba.Connection
        pFields = New SortedList
        pValues = New ArrayList
    End Sub

    <DebuggerStepThroughAttribute()> _
    Public Sub ConnectDatabase(ByVal ConnectionString As String)
        pc.ConnectionString = ConnectionString
        pc.Open()
    End Sub

    Public Property Connection() As OleDbConnection
        Get
            Return pc
        End Get
        Set(ByVal value As OleDbConnection)
            pc = value
            If pc.State = ConnectionState.Closed Then
                pc.Open()
            End If
        End Set
    End Property
    Public Property CommantType() As DBCommandTypes
        Get
            Return pCmdType
        End Get
        Set(ByVal value As DBCommandTypes)
            pCmdType = value
        End Set
    End Property
    Public Property Table()
        Get
            Return pTable
        End Get
        Set(ByVal value)
            pTable = value
        End Set
    End Property
    Public Property Fields(ByVal Name As String)
        Get
            Return pFields(Name)
        End Get
        Set(ByVal value)
            pFields(Name) = value
        End Set
    End Property
    Public ReadOnly Property Values()
        Get
            Return pValues
        End Get
    End Property
    Public Property Where() As String
        Get
            Return pWhere
        End Get
        Set(ByVal value As String)
            pWhere = value
        End Set
    End Property

    Public Sub SetFieldAndValues(ByVal FieldValuesPair As IDictionaryEnumerator)
        FieldValuesPair.Reset()

        While FieldValuesPair.MoveNext
            pFields.Add(FieldValuesPair.Key, FieldValuesPair.Value)
        End While

    End Sub

    Public Sub Reset()
        pCmdType = DBCommandTypes.NONE
        pTable = ""
        pFields.Clear()
        pValues.Clear()
        pWhere = ""

    End Sub
    Public Function Execute() As Object

        Dim sql As String
        Dim ret As Object

        Select Case pCmdType

            Case DBCommandTypes.SELECT_SCALAR
                sql = "SELECT  "
                sql &= "("
                sql &= "[" & pFields.Keys(0) & "]"
                For i As Integer = 1 To pFields.Count - 1
                    sql &= ", [" & pFields.Keys(i) & "]"
                Next
                sql &= ") "
                sql &= "FROM " & pTable

                sql &= " WHERE " & Where

                Dim cmd As New OleDbCommand(sql, Me.pc)
                ret = cmd.ExecuteScalar

            Case DBCommandTypes.INSERT

                sql = "INSERT INTO [" & pTable & "] "
                sql &= "("
                sql &= "[" & pFields.Keys(0) & "]"
                For i As Integer = 1 To pFields.Count - 1
                    sql &= ", [" & pFields.Keys(i) & "]"
                Next
                sql &= ") "
                sql &= "VALUES "
                sql &= "("
                sql &= QuoteValue(pFields.Values(0))
                For i As Integer = 1 To pFields.Count - 1
                    sql &= ", " & QuoteValue(pFields.Values(i))
                Next
                sql &= ")"

                Dim cmd As New OleDbCommand(sql, Me.pc)
                Try
                    ret = cmd.ExecuteNonQuery
                Catch ex As OleDbException
                    If IgnorePrimaryKeyViolation andalso ex.ErrorCode = -2147467259 Then
                        Exit Try
                    End If
                    Throw ex
                End Try


            Case DBCommandTypes.UPDATE

                sql = "UPDATE [" & pTable & "] "
                sql &= "SET "

                sql &= "[" & pFields.Keys(0) & "] = " & QuoteValue(pFields.Values(0))
                For i As Integer = 1 To pFields.Count - 1
                    sql &= ", [" & pFields.Keys(i) & "] = " & QuoteValue(pFields.Values(i))
                Next
                sql &= " WHERE " & Where

                Dim cmd As New OleDbCommand(sql, Me.pc)
                ret = cmd.ExecuteNonQuery

        End Select

        Return ret

    End Function

    Public Function GetSingleValue(ByVal sql As String, Optional ByVal FieldName As String = Nothing) As Object
        'se è richiesta una sola colonna l'array è singolo
        ' se no è un array di array
        Dim cmd As OleDbCommand
        Dim ret As Object
        Dim pos As Integer

        cmd = New OleDbCommand(sql, pc)
        Dim r As IDataReader = cmd.ExecuteReader

        pos = 0
        If Not FieldName Is Nothing Then
            pos = r.GetOrdinal(FieldName)
        End If

        If r.Read() Then
            ret = r.GetValue(pos)
        Else
            ret = Nothing
        End If

        r.Close()

        Return ret

    End Function
    Public Function GetValuesArray(ByVal sql As String) As ArrayList
        'se è richiesta una sola colonna l'array è singolo
        ' se no è un array di array
        Dim cmd As OleDbCommand
        Dim ret As ArrayList
        cmd = New OleDbCommand(sql, pc)
        Dim r As IDataReader = cmd.ExecuteReader
        If r.FieldCount > 1 Then
            Dim row(r.FieldCount - 1) As Object
            ret = New ArrayList
            While r.Read
                ret.Add(r.GetValues(row))
            End While
        Else
            ret = New ArrayList
            While r.Read
                ret.Add(r.GetValue(0))
            End While
        End If
        r.Close()

        Return ret

    End Function

    Public Function GetNamesValuesArray(ByVal sql As String) As Array
        'se è richiesta una sola colonna l'array è singolo
        ' se no è un array di array
        Dim cmd As OleDbCommand
        Dim ret As ArrayList
        cmd = New OleDbCommand(sql, pc)
        Dim r As IDataReader = cmd.ExecuteReader

        ret = New ArrayList
        While r.Read
            Dim row As New SortedList
            For i As Integer = 0 To r.FieldCount - 1
                row.Add(r.GetName(i), r.GetValue(i))
            Next
            ret.Add(row)
        End While

        r.Close()

        Return ret.ToArray(GetType(SortedList))

    End Function

    Private Function QuoteValue(ByVal Value As Object) As String
        Dim ret As String
        If Value Is Nothing Then
            Return "''"
        End If
        If TypeOf Value Is String Then
            'Bonifico la stringa
            'se contiene degli " li faccio diventare "" 
            'se contiene degli ' li faccio diventare ''
            ret = Value
            ret = ret.Replace("'", "''")
            ret = "'" & ret & "'"

        ElseIf TypeOf Value Is Integer Then
            ret = Value

        ElseIf TypeOf Value Is Double Then
            ret = Format(Value)
            ret = ret.Replace(",", ".")

        ElseIf TypeOf Value Is Date Then
            ret = "#" & Format(Value, "MM-dd-yyyy HH:mm:ss") & "#"

        Else
            'Stringa
            ret = Value
            ret = ret.Replace("'", "''")
            ret = "'" & ret & "'"

        End If
        Return ret
    End Function

End Class

Public Enum DBCommandTypes
    [NONE]
    [SELECT_SCALAR]
    [SELECT]
    [UPDATE]
    [INSERT]
    [DELETE]
End Enum