Attribute VB_Name = "modTimeAndDateHandling"
Option Explicit

Public Function GetTodaysDate(ByRef strDelimiter As String) As String
    Dim datDate As Date
    
    datDate = Date
    
    GetTodaysDate = Format(datDate, "yyyy") & strDelimiter & _
        Format(datDate, "mm") & strDelimiter & _
        Format(datDate, "dd")
End Function

Public Function GetActualTime(ByRef strDelimiter As String) As String
    Dim strTime As String
    
    strTime = Time
    
    GetActualTime = Format(strTime, "HH") & strDelimiter & _
        Format(strTime, "mm") & strDelimiter & _
        Format(strTime, "ss")
End Function
