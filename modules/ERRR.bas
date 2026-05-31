Attribute VB_Name = "ERRR"
Option Compare Database
Option Explicit
 Function AccessAndJetErrorsTable() As Boolean
      Dim dbs As Database, tdf As TableDef, fld As Field
      Dim rst As New ADODB.Recordset, lngCode As Long
      Dim strAccessErr As String
      Const conAppObjectError = "Application-defined or object-defined error"

       On Error GoTo Error_AccessAndJetErrorsTable
      ' Create Errors table with ErrorNUMBER and ErrorDescription fields.
      Set dbs = CurrentDb
      Set tdf = dbs.CreateTableDef("AccessAndJetErrors")
      Set fld = tdf.CreateField("ErrorCode", dbLong)

      tdf.Fields.Append fld
      Set fld = tdf.CreateField("ErrorString", dbMemo)
      tdf.Fields.Append fld

      dbs.TableDefs.Append tdf
      ' Open recordset  On Errors table.
      Set rst = dbs.OpenRecordset("AccessAndJetErrors")
      ' Loop through error codes.
      For lngCode = 0 To 3500
                             On Error Resume Next
          ' Raise each error.
          strAccessErr = AccessError(lngCode)
          DoCmd.Hourglass True
          ' Skip error NUMBERs without associated strings.
          If strAccessErr <> "" Then

              ' Skip codes that generate application or object-defined errors.
              If strAccessErr <> conAppObjectError Then
                  ' Add each error code and string to Errors table.
                  rst.AddNew
                  rst!ErrorCode = lngCode
                  ' Append string to memo field.
                  rst!ErrorString.AppendChunk strAccessErr
                  rst.update
              End If
          End If
      Next lngCode
      ' Close recordset.
      rst.Close
      DoCmd.Hourglass False
      RefreshDatabaseWindow
      MsgBox "Access and Jet errors table created."

      AccessAndJetErrorsTable = True

Exit_AccessAndJetErrorsTable:
      Exit Function

Error_AccessAndJetErrorsTable:
      MsgBox Err & ": " & Err.Description
      AccessAndJetErrorsTable = False
      Resume Exit_AccessAndJetErrorsTable
  End Function
