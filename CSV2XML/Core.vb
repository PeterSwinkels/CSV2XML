'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module CoreModule
   Private Const DELIMITER As Char = ";"c   'Defines the delimiter used to separate columns.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Dim DataName As String = Nothing
      Dim InputData As List(Of String) = Nothing
      Dim InputFile As String = If(GetCommandLineArgs.Count > 1, GetCommandLineArgs.Last, Nothing)
      Dim OutputFile As String = Nothing
      Dim Table As DataTable = Nothing

      If Not InputFile = Nothing Then
         DataName = Path.GetFileNameWithoutExtension(InputFile)
         InputData = New List(Of String)(File.ReadAllLines(InputFile))
         OutputFile = $"{Path.Combine(Path.GetDirectoryName(InputFile), DataName)}.xml"

         Table = New DataTable() With {.TableName = $"{DataName}Table"}

         For Each Column As String In InputData.First.Split(DELIMITER)
            Table.Columns.Add(New DataColumn With {.Caption = Column, .ColumnName = Column, .DataType = GetType(String)})
         Next Column

         InputData.RemoveAt(0)

         For Each Line As String In InputData
            Table.Rows.Add(Line.Split(DELIMITER))
         Next Line

         Using DataSetO As New DataSet() With {.DataSetName = $"{DataName}DataSet"}
            DataSetO.Tables.Add(Table)
            DataSetO.WriteXml(OutputFile)
         End Using
      End If
   End Sub
End Module
