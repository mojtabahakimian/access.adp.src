Attribute VB_Name = "mdlDispatchToAustraliaPostHandshake"

Option Explicit
Option Compare Text


' Constant server name for date/time stamp,
' vbNullString uses local workstation date/time.


' Structure of return type for function GetTimeFromServer()
Public Type ServerInfo
    SupplierDate As Date
    SupplierName As String
End Type


' Structure for Time of Day.
Public Type TIME_OF_DAY_INFO
    t_elapsedt As Long
    t_msecs As Long
    t_hours As Long
    t_mins As Long
    t_secs As Long
    t_hunds As Long
    t_timezone As Long
    t_tinterval As Long
    t_day As Long
    t_month As Long
    t_year As Long
    t_weekday As Long
End Type
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Private Declare Function NetApiBufferFree Lib "NETAPI32.DLL" (ByVal bufptr As Long) As Long

Private Declare Function NetRemoteTOD Lib "NETAPI32.DLL" (ByVal server As Long, _
                                                          ByRef bufptr As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDest As Any, _
                                                  ByRef pSrc As Any, _
                                                  ByVal ByteLen As Long)

Public Declare Function GetComputerName Lib "kernel32" Alias _
                       "GetComputerNameA" (ByVal lpBuffer As String, _
                                           ByRef nSize As Long) As Long

Public Declare Function timeGetTime Lib "Winmm.dll" () As Long
'
Public Function GetdateFromServer() As Date
                    On Error Resume Next
    Dim DTT As Date, rst As New ADODB.Recordset
    rst.Open "select getdate()", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GetdateFromServer = rst.Fields(0)
    Else
        GetdateFromServer = VBA.Now()
    End If
End Function

Public Function GetdateFromServer1() As Date
                    On Error Resume Next
    Dim tdi As TIME_OF_DAY_INFO
    Dim lngBufptr As Long
    Dim DTT As Date
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    Forms![BUN].Form.Refresh
    If Forms![BASEKNOW]![SERVERNAM] = vbNullString Then
    Forms![BUN].Form.Refresh
        GetdateFromServer1 = VBA.Now()
    Else
    Forms![BUN].Form.Refresh
        If NetRemoteTOD(StrPtr(Forms![BASEKNOW]![SERVERNAM]), lngBufptr) = 0 Then
     Forms![BUN].Form.Refresh
           Call RtlMoveMemory(tdi, ByVal lngBufptr, Len(tdi))
    Forms![BUN].Form.Refresh
            GetdateFromServer1 = VBA.DateSerial(tdi.t_year, tdi.t_month, tdi.t_day) + VBA.TimeSerial(tdi.t_hours, tdi.t_mins, tdi.t_secs) - (tdi.t_timezone / 60 / 24)
    Forms![BUN].Form.Refresh
            Call NetApiBufferFree(lngBufptr)
    Forms![BUN].Form.Refresh
        Else
    Forms![BUN].Form.Refresh
            GetdateFromServer1 = VBA.Now()
     Forms![BUN].Form.Refresh
       End If
    Forms![BUN].Form.Refresh
    End If
End Function
