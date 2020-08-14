'***********************************************************************
'Macros by Eric Bentzen, 22th August 2015.
'How to make a calendar for picking dates with VBA only - no ActiveX
'and lack of compatibility between versions.
'You can format the selected dates in the two userform
'procedures: "FillFirstDay" and "FillSecondDay".

'Bug fix December 2015: Dates were not formatted correctly,
'if the date format in the system settings was MM/DD/Year
'This is now fixed by the userform's ReturnDate function.

'You can find more VBA and macro stuff at:
'http://sitestory.dk/excel_vba/vba-start-page.htm
'***********************************************************************
Option Explicit
Public colLabelEvent As Collection 'Collection of labels for event handling
Public colLabels As Collection     'Collection of the date labels
Public bSecondDate As Boolean      'True if finding second date
Public sActiveDay As String        'Last day selected
Public lDays As Long               'Number of days in month
Public lFirstDay As Long           'Day selected, e.g. 19th
Public lStartPos As Long
Public lSelMonth As Long           'The selected month
Public lSelYear As Long            'The selected year
Public lSelMonth1 As Long          'Used to check if same date is selected twice
Public lSelYear1 As Long           'Used to check if same date is selected twice


Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public strSQL As String

Public g_objFSO As Scripting.FileSystemObject
Public g_scrText As Scripting.TextStream

Public Sub OpenDB()
    If cnn.State = adStateOpen Then cnn.Close
    'On Error GoTo Handle
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & _
        [DB_PATH]
        cnn.Open
    Exit Sub
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, "Mulo2", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub closeRS()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If cnn.State = adStateOpen Then cnn.Close
    Set cnn = Nothing
End Sub

Public Sub logError(conn As ADODB.Connection, hoja As String, func As String)
    If conn.Errors.count > 0 Then
        Call Error_Handle(conn.Errors.Item(0).Source, hoja & " - " & func, strSQL, conn.Errors.Item(0).Number, conn.Errors.Item(0).Description)
        conn.Errors.Clear
        closeRS
    End If
End Sub

Public Function LogFile_WriteError(ByVal sRoutineName As String, _
                             ByVal sMessage As String)
Dim sText As String
   On Error GoTo ErrorHandler
   If (g_objFSO Is Nothing) Then
      Set g_objFSO = New FileSystemObject
   End If
   If (g_scrText Is Nothing) Then
      If (g_objFSO.FileExists([LOG_PATH]) = False) Then
         Set g_scrText = g_objFSO.OpenTextFile([LOG_PATH], IOMode.ForWriting, True)
      Else
         Set g_scrText = g_objFSO.OpenTextFile([LOG_PATH], IOMode.ForAppending)
      End If
   End If
   sText = sText & Format(Date, "DD/MM/YYYY") & " " & Time() & "|"
   sText = sText & sRoutineName & "|"
   sText = sText & sMessage & "|"
   g_scrText.WriteLine sText
   g_scrText.Close
   Set g_scrText = Nothing
   Exit Function
ErrorHandler:
   Set g_scrText = Nothing
   Call MsgBox("No se pudo escribir en el fichero log", vbCritical, "LogFile_WriteError")
End Function

Public Sub Error_Handle(ByVal sRoutineName As String, _
                         ByVal sObject As String, _
                         ByVal currentStrSQL As String, _
                         ByVal sErrorNo As String, _
                         ByVal sErrorDescription As String)
Dim sMessage As String
   sMessage = sObject & "|" & currentStrSQL & "|" & sErrorNo & "|" & sErrorDescription & "|" & Application.UserName
   Call MsgBox(sErrorNo & vbCrLf & sErrorDescription, vbCritical, sRoutineName & " - " & sObject & " - Error")
   Call LogFile_WriteError(sRoutineName, sMessage)
End Sub

Public Function fechaStrStr(fechaDDMMYYYY As String)
    Dim splitfecha As Variant
    splitfecha = Split(fechaDDMMYYYY, "/")
    fechaStrStr = splitfecha(2) & "-" & splitfecha(1) & "-" & splitfecha(0)
End Function

Public Function fechaDateStr(fechaDate As Date)
    fechaDateStr = Format(fechaDate, "YYYY") & "-" & Format(fechaDate, "MM") & "-" & Format(fechaDate, "DD")
End Function
