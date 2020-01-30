# vba
    Sub CheckDeclines(oRequest As MeetingItem)
    ' only check declines
     If oRequest.MessageClass <> "IPM.Schedule.Meeting.Resp.Neg" Then
      Exit Sub
    End If

    Dim oAppt As AppointmentItem
    Set oAppt = oRequest.GetAssociatedAppointment(True)
    Dim myAttendees As Integer
    Dim objAttendees As Outlook.recipients
    Set objAttendees = oAppt.recipients

    For x = 1 To objAttendees.Count
       If objAttendees(x).Type = olRequired Then
         oAppt.Categories = "Declined;"
         objAttendees.Count = myAttendees
       End If
    Next

    If myAttendees = 1 Then
    oAppt.Categories = "Declined;"
    oAppt.Categories = "Yellow Category"
    Else
    If myAttendees = 2 Then
    oAppt.Categories = "Declined;"
    oAppt.Categories = "Orange Category"
    Else
    oAppt.Categories = "Declined;"
    oAppt.Categories = "Red Category"
    End If
    End If

    oAppt.Display

    End Sub
