Private Sub Document_ContentControlOnExit(ByVal ContentControl _
As ContentControl, Cancel As Boolean)
    
    insertCountDown (ContentControl.Range.Text)
End Sub


Function insertCountDown(strEndDate As String)

   Dim objRangeDate
   Dim objDoc
    Dim i As Integer
    Dim j As Integer
    Dim leftDays As Integer
    
    Dim objDate As Date
    Dim startDate As Date
    Dim endDate As Date
    
    Dim strDate As String


   Set objselection = Selection
   startDate = Now
   
   
   
   If Not IsDate(strEndDate) Then
    MsgBox ("data non valida!!")
    Exit Function
   End If
   
   endDate = Format(strEndDate, "dd/mm/yyyy")
 
   leftDays = Work_Days(startDate, endDate)
   i = leftDays
   j = 0

ActiveDocument.Content.InsertBreak

Do
   
      objDate = DateAdd("d", j, startDate)
   
    If Not IsWeekend(objDate) Then
        
       
       
       strDate = Format(objDate, "dddd dd mmmm yyyy")
       
        
        With objselection.Font
                 .Bold = False
                .Name = "Arial"
              .Size = 52
        End With
        objselection.TypeText Text:=strDate & vbCrLf
        objselection.InsertBreak Type:=wdSectionBreakContinuous
             
        
        With objselection.Font
                    .Bold = True
                    .Name = "Arial"
                    .Size = 490
        End With
        objselection.TypeText Text:="-" & i & vbCrLf
        objselection.InsertBreak Type:=wdSectionBreakContinuous
        
        
        'diminuisco il countdown solo quando la data non corrisponde al weekend
        i = i - 1
    End If
    
        'numero progessivo per calcolare la data
        j = j + 1
    
Loop Until i < 0

   
ActiveDocument.Paragraphs.Format.Alignment = _
 wdAlignParagraphCenter
   
End Function
 
 
 Function Work_Days(BegDate As Variant, endDate As Variant) As Integer
 
 Dim WholeWeeks As Variant
 Dim DateCnt As Variant
 Dim EndDays As Integer
 
 On Error GoTo Err_Work_Days
 
 BegDate = DateValue(BegDate)
 endDate = DateValue(endDate)
 WholeWeeks = DateDiff("w", BegDate, endDate)
 DateCnt = DateAdd("ww", WholeWeeks, BegDate)
 EndDays = 0
 
 Do While DateCnt <= endDate
 
 If Not IsWeekend(DateValue(DateCnt)) Then
 EndDays = EndDays + 1
 End If

 DateCnt = DateAdd("d", 1, DateCnt)
 Loop
 
 Work_Days = WholeWeeks * 5 + EndDays
 
Exit Function
 
Err_Work_Days:
 
 ' If either BegDate or EndDate is Null, return a zero
 ' to indicate that no workdays passed between the two dates.
 
 If Err.Number = 94 Then
 Work_Days = 0
 Exit Function
 Else
' If some other error occurs, provide a message.
 MsgBox "Error " & Err.Number & ": " & Err.Description
 End If
 
End Function


Private Function IsWeekend(dtmTemp As Date) As Boolean
    ' If your weekends aren't Saturday (day 7)
    ' and Sunday (day 1), change this routine
    ' to return True for whatever days
    ' you DO treat as weekend days.
    Select Case Weekday(dtmTemp)
        Case vbSaturday, vbSunday
            IsWeekend = True
        Case Else
            IsWeekend = False
    End Select
End Function



