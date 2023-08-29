Attribute VB_Name = "Module0Dates"
Option Explicit

Sub showDateInfo()

      Dim dateStart As Date
      Dim dateStop As Date
      Dim indexMonth As Integer
      Dim numDaysInMonth As Integer
      Dim nameMonth As String
      
            'Ask user for starting date (Assumption is that this will be the 1st of the month.)
            'Determine needed information from starting date
            dateStart = InputBox(prompt:= _
                    "Please give the starting date of the month in the form mm/dd/yyyy")
            dateStop = WorksheetFunction.EoMonth(dateStart, 0)
            MsgBox prompt:="The month starts on " & dateStart & _
                    " and ends on " & dateStop
                    
            indexMonth = Month(dateStart)
            nameMonth = MonthName(indexMonth)
            MsgBox prompt:="The month you gave is " & nameMonth & _
                    ", it is the month number " & indexMonth & " of the year."
                    
            numDaysInMonth = dateStop - dateStart + 1
            MsgBox prompt:="The number of days in the month you gave is " & _
                    numDaysInMonth
        
End Sub


