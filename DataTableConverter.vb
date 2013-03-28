Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Data

Namespace Serialization

    ''' <summary>
    ''' Converts a DataTable to JSON. Note no support for deserialization
    ''' </summary>
    Public Class DataTableConverter
        Inherits JsonConverter

        ''' <summary>
        ''' Determines whether this instance can convert the specified object type.
        ''' </summary>
        ''' <param name="objectType">Type of the object.</param><returns>
        '''   <c>true</c> if this instance can convert the specified object type; otherwise, <c>false</c>.
        ''' </returns>
        Public Overrides Function CanConvert(objectType As System.Type) As Boolean
            'Return objectType = GetType(DataTable)
            Return GetType(DataTable).IsAssignableFrom(objectType)
        End Function

        ''' <summary>
        ''' Reads the json.
        ''' </summary>
        ''' <param name="reader">The reader.</param>
        ''' <param name="objectType">Type of the object.</param>
        ''' <param name="existingValue">The existing value.</param>
        ''' <param name="serializer">The serializer.</param><returns></returns>
        Public Overrides Function ReadJson(reader As Newtonsoft.Json.JsonReader, objectType As System.Type, existingValue As Object, serializer As Newtonsoft.Json.JsonSerializer) As Object
            Dim jObject As JObject = jObject.Load(reader)

            Dim table As DataTable = New DataTable()

            If jObject("TableName") IsNot Nothing Then
                table.TableName = jObject("TableName").ToString()
            End If

            If jObject("Columns") Is Nothing Then Return table

            ' Loop through the columns in the table and apply any properties provided
            For Each jColumn As JObject In jObject("Columns")
                Dim column As New DataColumn
                Dim token As JToken

                token = jColumn.SelectToken("AllowDBNull")
                If token IsNot Nothing Then
                    column.AllowDBNull = token.Value(Of Boolean)()
                End If

                token = jColumn.SelectToken("AutoIncrement")
                If token IsNot Nothing Then
                    column.AutoIncrement = token.Value(Of Boolean)()
                End If

                token = jColumn.SelectToken("AutoIncrementSeed")
                If token IsNot Nothing Then
                    column.AutoIncrementSeed = token.Value(Of Long)()
                End If

                token = jColumn.SelectToken("AutoIncrementStep")
                If token IsNot Nothing Then
                    column.AutoIncrementStep = token.Value(Of Long)()
                End If

                token = jColumn.SelectToken("Caption")
                If token IsNot Nothing Then
                    column.Caption = token.Value(Of String)()
                End If

                token = jColumn.SelectToken("ColumnName")
                If token IsNot Nothing Then
                    column.ColumnName = token.Value(Of String)()
                End If

                ' Allowed data types: http://msdn.microsoft.com/en-us/library/system.data.datacolumn.datatype.aspx
                token = jColumn.SelectToken("DataType")
                If token IsNot Nothing Then
                    Dim dataType As String = token.Value(Of String)()
                    If dataType = "Byte[]" Then
                        column.DataType = GetType(System.Byte())
                    Else
                        ' All allowed data types exist in the System namespace
                        column.DataType = Type.GetType(String.Concat("System.", dataType))
                    End If
                End If

                token = jColumn.SelectToken("DateTimeMode")
                If token IsNot Nothing Then
                    column.DateTimeMode = [Enum].Parse(GetType(System.Data.DataSetDateTime), token.Value(Of String)())
                End If

                If Not column.AutoIncrement Then ' Can't set default value on auto increment column
                    token = jColumn.SelectToken("DefaultValue")
                    If token IsNot Nothing Then
                        ' If a default value is provided then cast to the columns data type
                        Select Case column.DataType
                            Case GetType(System.Boolean)
                                Dim defaultValue As Boolean
                                If Boolean.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Byte)
                                Dim defaultValue As Byte
                                If Byte.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Char)
                                Dim defaultValue As Char
                                If Char.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.DateTime)
                                Dim defaultValue As DateTime
                                If DateTime.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Decimal)
                                Dim defaultValue As Decimal
                                If Decimal.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Double)
                                Dim defaultValue As Double
                                If Double.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Guid)
                                Dim defaultValue As Guid
                                If Guid.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Int16)
                                Dim defaultValue As Int16
                                If Int16.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Int32)
                                Dim defaultValue As Int32
                                If Int32.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Int64)
                                Dim defaultValue As Int64
                                If Int64.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.SByte)
                                Dim defaultValue As SByte
                                If SByte.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Single)
                                Dim defaultValue As Single
                                If Single.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.String)
                                column.DefaultValue = token.ToString()
                            Case GetType(System.TimeSpan)
                                Dim defaultValue As TimeSpan
                                If TimeSpan.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.UInt16)
                                Dim defaultValue As UInt16
                                If UInt16.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.UInt32)
                                Dim defaultValue As UInt32
                                If UInt32.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.UInt64)
                                Dim defaultValue As UInt64
                                If UInt64.TryParse(token.ToString(), defaultValue) Then
                                    column.DefaultValue = defaultValue
                                End If
                            Case GetType(System.Byte())
                                ' Leave as null
                        End Select
                    End If
                End If

                token = jColumn.SelectToken("MaxLength")
                If token IsNot Nothing Then
                    column.MaxLength = token.Value(Of Integer)()
                End If

                token = jColumn.SelectToken("ReadOnly")
                If token IsNot Nothing Then
                    column.ReadOnly = token.Value(Of Boolean)()
                End If

                token = jColumn.SelectToken("Unique")
                If token IsNot Nothing Then
                    column.Unique = token.Value(Of Boolean)()
                End If

                table.Columns.Add(column)
            Next

            ' Add the rows to the table
            If jObject("Rows") IsNot Nothing Then
                For Each jRow As JArray In jObject("Rows")
                    Dim row As DataRow = table.NewRow()
                    ' Each row is just an array of objects
                    row.ItemArray = jRow.ToObject(Of System.Object())()
                    table.Rows.Add(row)
                Next
            End If

            ' Add the primary key to the table if supplied
            If jObject("PrimaryKey") IsNot Nothing Then
                Dim primaryKey As New List(Of DataColumn)
                For Each jPrimaryKey As JValue In jObject("PrimaryKey")
                    Dim column As DataColumn = table.Columns(jPrimaryKey.ToString())
                    If column Is Nothing Then
                        Throw New ApplicationException("Invalid primary key.")
                    Else
                        primaryKey.Add(column)
                    End If
                Next
                table.PrimaryKey = primaryKey.ToArray()
            End If

            Return table
        End Function

        ''' <summary>
        ''' Writes the json.
        ''' </summary>
        ''' <param name="writer">The writer.</param>
        ''' <param name="value">The value.</param>
        ''' <param name="serializer">The serializer.</param>
        Public Overrides Sub WriteJson(writer As Newtonsoft.Json.JsonWriter, value As Object, serializer As Newtonsoft.Json.JsonSerializer)
            Dim table As DataTable = TryCast(value, DataTable)

            writer.WriteStartObject()

            writer.WritePropertyName("TableName")
            writer.WriteValue(table.TableName)

            writer.WritePropertyName("Columns")
            writer.WriteStartArray()

            For Each column As DataColumn In table.Columns
                writer.WriteStartObject()

                writer.WritePropertyName("AllowDBNull")
                writer.WriteValue(column.AllowDBNull)
                writer.WritePropertyName("AutoIncrement")
                writer.WriteValue(column.AutoIncrement)
                writer.WritePropertyName("AutoIncrementSeed")
                writer.WriteValue(column.AutoIncrementSeed)
                writer.WritePropertyName("AutoIncrementStep")
                writer.WriteValue(column.AutoIncrementStep)
                writer.WritePropertyName("Caption")
                writer.WriteValue(column.Caption)
                writer.WritePropertyName("ColumnName")
                writer.WriteValue(column.ColumnName)
                writer.WritePropertyName("DataType")
                writer.WriteValue(column.DataType.Name)
                writer.WritePropertyName("DateTimeMode")
                writer.WriteValue(column.DateTimeMode.ToString())
                writer.WritePropertyName("DefaultValue")
                writer.WriteValue(column.DefaultValue)
                writer.WritePropertyName("MaxLength")
                writer.WriteValue(column.MaxLength)
                writer.WritePropertyName("Ordinal")
                writer.WriteValue(column.Ordinal)
                writer.WritePropertyName("ReadOnly")
                writer.WriteValue(column.ReadOnly)
                writer.WritePropertyName("Unique")
                writer.WriteValue(column.Unique)

                writer.WriteEndObject()
            Next

            writer.WriteEndArray()

            writer.WritePropertyName("Rows")
            writer.WriteStartArray()

            For Each row As DataRow In table.Rows
                If row.RowState <> DataRowState.Deleted AndAlso row.RowState <> DataRowState.Detached Then
                    writer.WriteStartArray()

                    For index As Integer = 0 To table.Columns.Count - 1
                        writer.WriteValue(row(index))
                    Next

                    writer.WriteEndArray()
                End If
            Next

            writer.WriteEndArray()

            ' Write out primary key if the table has one. This will be useful when deserializing the table.
            ' We will write it out as an array of column names
            writer.WritePropertyName("PrimaryKey")
            writer.WriteStartArray()
            If table.PrimaryKey.Length > 0 Then
                For Each column As DataColumn In table.PrimaryKey
                    writer.WriteValue(column.ColumnName)
                Next
            End If
            writer.WriteEndArray()

            writer.WriteEndObject()
        End Sub

    End Class

End Namespace
