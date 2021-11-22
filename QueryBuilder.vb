Imports System.Text

Public Class QueryBuilder
    Public Shared Function AppendPagginationToQuery(sql As String, pageNo As Integer, pageSize As Integer, Optional orderBy As String = "1") As String

        sql += $" Order by {orderBy} Offset({pageNo} - 1)* {pageSize} rows"
        sql += $" Fetch Next {pageSize} rows only "

        Return sql
    End Function

    Public Shared Function TotalCount(table As String, Optional condition As String = "") As String

        Dim query As String
        query = $" select Count(*) from {table} "
        If Not String.IsNullOrEmpty(condition) Then
            query += $" where {condition}"
        End If

        Return query
    End Function


    Public Shared Function SelectQuery(ByVal fields As List(Of String), table As String, Optional condition As String = "") As String

        Dim selectedFields As String = ""
        Dim isFirstItem As Boolean = True
        For Each field In fields
            If field = "*" Then
                selectedFields = "*"
                Exit For

            End If
            If isFirstItem Then
                selectedFields = field
                isFirstItem = False

            Else
                selectedFields += ("," + field)
            End If

        Next

        Dim query As String
        query = $" select {selectedFields} from {table} "

        If Not String.IsNullOrEmpty(condition) Then
            query += $" where {condition}"

        End If

        Return query
    End Function


    Public Shared Function SelectQuery(ByVal fields As List(Of String), table As String, pageNo As Integer, pageSize As Integer, Optional orderBy As String = "1", Optional condition As String = "") As String

        Dim selectedFields As String = ""
        Dim isFirstItem As Boolean = True
        For Each field In fields
            If field = "*" Then
                selectedFields = "*"
                Exit For

            End If
            If isFirstItem Then
                selectedFields = field
                isFirstItem = False

            Else
                selectedFields += ("," + field)
            End If

        Next

        Dim query As String
        query = $" select {selectedFields} from {table} "
        If Not String.IsNullOrEmpty(condition) Then
            query += $" where {condition}"

        End If
        query = AppendPagginationToQuery(query, pageNo, pageSize, orderBy)

        Return query
    End Function




    Public Shared Function SelectQuery(ByVal fields As List(Of String), table As String, innerjoins As List(Of KeyValuePair(Of String, String)), Optional pageNo As Integer = 0, Optional pageSize As Integer = 0, Optional orderBy As String = "1", Optional condition As String = "") As String

        Dim selectedFields As String = ""
        Dim isFirstItem As Boolean = True
        For Each field In fields
            If field = "*" Then
                selectedFields = "*"
                Exit For

            End If
            If isFirstItem Then
                selectedFields = field
                isFirstItem = False

            Else
                selectedFields += ("," + field)
            End If

        Next

        '' Configure JOINS
        Dim joins As String = ""
        For Each iJoin In innerjoins
            joins += $" full outer join {iJoin.Key} on {iJoin.Key}.Id = {table}.{iJoin.Value} "
        Next


        Dim query As String
        query = $" select {selectedFields} from {table} "
        query += joins

        If Not String.IsNullOrEmpty(condition) Then
            query += $" where {condition}"
        End If


        If pageNo > 0 And pageSize > 0 Then
            query = AppendPagginationToQuery(query, pageNo, pageSize, "1")
        End If


        Return query
    End Function


    'Commands 

    Public Shared Function Update(ByVal fields As List(Of KeyValuePair(Of String, String)), table As String, Optional condition As String = "") As String


        Dim itemsToUpdate As String = ""
        Dim isFirstItem As Boolean = True


        For Each item In fields

            If isFirstItem Then
                itemsToUpdate = $"[{item.Key}] = {item.Value}"
                isFirstItem = False
            Else
                itemsToUpdate += ("," + $"[{item.Key}] = {item.Value}")
            End If




        Next




        Dim query As String = $"UPDATE {table} SET {itemsToUpdate} "
        If Not String.IsNullOrEmpty(condition) Then
            query += $" where {condition}"
        End If

        Return query

    End Function



    Public Shared Function Insert(ByVal fields As List(Of KeyValuePair(Of String, String)), table As String) As String


        Dim query As New StringBuilder

        Try

            query.Append($"INSERT INTO {table} ( ")
            Dim isFirstItem As Boolean = True
            For Each item In fields

                If isFirstItem Then
                    query.Append(item.Key)
                    isFirstItem = False
                Else
                    query.Append(", " + item.Key)
                End If
            Next
            query.Append(" ) VALUES ( ")

            isFirstItem = True
            For Each item In fields

                If isFirstItem Then
                    query.Append(item.Value)
                    isFirstItem = False
                Else
                    query.Append(", " + item.Value)
                End If
            Next

            query.Append(" )  ")

        Catch ex As Exception
            Return ""
        End Try

        Return query.ToString()

    End Function


End Class
