'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics.FileVersionInfo
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Reflection

'This module contains this program's core procedures.
Public Module CoreModule
   Private Const DELIMITER As Char = ";"c                                                                                            'Defines the delimiter used to separate columns.
   Private ReadOnly EXECUTABLE_NAME As String = Path.GetFileName(GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileName)  'Defines this program's executable file name.

   'This procedure displays any errors that occur.
   Private Sub DisplayError(ExceptionO As Exception)
      Try
         With Console.Error
            .WriteLine($"Error: {ExceptionO.Message}")
            .WriteLine("Press Enter to continue...")
            .WriteLine()
         End With

         Console.ReadLine()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim DataName As String = Nothing
         Dim InputData As List(Of String) = Nothing
         Dim InputFile As String = If(GetCommandLineArgs.Count > 1, GetCommandLineArgs.Last, Nothing)
         Dim OutputFile As String = Nothing
         Dim Table As DataTable = Nothing

         If InputFile = Nothing Then
            Console.WriteLine($"{ProgramInformation()}{NewLine}")
            Console.WriteLine($"{My.Application.Info.Description}{NewLine}")
            Console.WriteLine($"Usage: {EXECUTABLE_NAME} INPUT_FILE{NewLine}")
         Else
            DataName = Path.GetFileNameWithoutExtension(InputFile)
            InputData = New List(Of String)(File.ReadAllLines(InputFile))
            OutputFile = $"{Path.Combine(Path.GetDirectoryName(InputFile), DataName)}.xml"

            Console.WriteLine($"""{InputFile}"" -> ""{OutputFile}""")

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

            Console.WriteLine($"{NewLine}Done.")
         End If
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure returns information about this program.
   Private Function ProgramInformation() As String
      Try
         Dim Information As String = Nothing

         With My.Application.Info
            Information = $"{ .Title} v{ .Version} - by: { .CompanyName}, ***{ .Copyright}***"
         End With

         Return Information
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
