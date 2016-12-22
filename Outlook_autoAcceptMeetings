Sub AutoAcceptMeetings(oRequest As MeetingItem)
    'If not a meeting request, leave the sub
    If oRequest.MessageClass <> "IPM.Schedule.Meeting.Request" Then
        Exit Sub
    End If

    'Working on the appointment
    Dim receivedAppt As AppointmentItem
    Set receivedAppt = oRequest.GetAssociatedAppointment(True)

    'Accepting the meeting
    Dim oResponse
    Set oResponse = receivedAppt.Respond(olMeetingAccepted, True)
    
    receivedAppt.BusyStatus = olOutOfOffice
   
    'create an additional meetings traveling, 2hrs before, 1hr after
    Dim frontStubAppt As AppointmentItem
    Set frontStubAppt = Application.CreateItem(olAppointmentItem)

    'DateAdd: interval, number, date; n=minutes
    frontStubAppt.StartInStartTimeZone = DateAdd("n", -120, receivedAppt.StartInStartTimeZone)
    frontStubAppt.EndInEndTimeZone = DateAdd("n", 0, receivedAppt.StartInStartTimeZone)
    frontStubAppt.BusyStatus = olOutOfOffice
    frontStubAppt.Subject = "Transit for travel (" + receivedAppt.Subject + ")"
    frontStubAppt.Save
    frontStubAppt.Close (olSave)
    
    Dim endStubAppt As AppointmentItem
    Set endStubAppt = Application.CreateItem(olAppointmentItem)
    endStubAppt.StartInStartTimeZone = DateAdd("n", 0, receivedAppt.EndInEndTimeZone)
    endStubAppt.EndInEndTimeZone = DateAdd("n", 60, receivedAppt.EndInEndTimeZone)
    endStubAppt.BusyStatus = olOutOfOffice
    endStubAppt.Subject = "Transit for travel (" + receivedAppt.Subject + ")"
    endStubAppt.Save
    endStubAppt.Close (olSave)

    receivedAppt.Save
    receivedAppt.Close (olSave)
    oRequest.Delete
End Sub
