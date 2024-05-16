# Outlook Meeting Attendance Reminder

This repository contains a VBA macro for Microsoft Outlook that helps send reminder emails to meeting attendees based on their response status. The macro can be used to send emails to attendees who have not responded ("None") or who have tentatively accepted the meeting ("Tentative").

## Features

- Automatically generate and send reminder emails to specific attendees.
- Customize email subject and body.
- Avoid sending duplicate emails to the meeting organizer.
- Easy integration with Outlook Ribbon for quick access.

## Installation

1. **Download the VBA code:**
   - Download the `MeetingAttendanceReminder.cls` file from the repository.

2. **Open Outlook Visual Basic for Applications (VBA) Editor:**
   - Open Outlook.
   - Press `Alt + F11` to open the VBA Editor.

3. **Import the VBA code into `ThisOutlookSession`:**
   - In the Project Explorer, find and double-click on `ThisOutlookSession` under `Project1 (VbaProject.otm)`.
   - Copy the code from `MeetingAttendanceReminder.cls` and paste it into the `ThisOutlookSession` module.
   - Alternatively, you can directly copy the provided code from the repository's README and paste it into `ThisOutlookSession`.

4. **Customize the Outlook Ribbon:**
   - Download the `MeetingAttendanceReminderRibbon.exportedUI` file from the repository.
   - Go to `File > Options > Customize Ribbon`.
   - Click on "Import/Export" and select "Import customization file".
   - Choose the `MeetingAttendanceReminderRibbon.exportedUI` file to add a button in the Outlook Ribbon for easy access to the macro.

## Usage

1. Highlight a meeting in your Outlook calendar.
2. Click on the "Send Reminder Email" button in the custom group on the Ribbon.
3. Choose whether to include attendees with "None" or "Tentative" response statuses.
4. The macro will generate and display an email with the appropriate recipients in the BCC field.

## Files

- `MeetingAttendanceReminder.cls`: The VBA macro code.
- `MeetingAttendanceReminderRibbon.exportedUI`: The exported UI customization file for the Outlook Ribbon.

## Version

- Current version: `0.1-beta`

## Contributing

Feel free to submit issues, fork the repository, and make pull requests. Contributions are welcome!

## License

This project is licensed under the MIT License.


## VBA Macro Code

Copy the following VBA code into the `ThisOutlookSession` module in Outlook VBA Editor:

```vba
Option Explicit

Sub SendEmailToSelectedResponders()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.Folder
    Dim olAppointment As Outlook.AppointmentItem
    Dim olRecipient As Outlook.Recipient
    Dim olMail As Outlook.MailItem
    Dim bccList As String
    Dim subject As String
    Dim emailBody As String
    Dim includeNone As VbMsgBoxResult
    Dim includeTentative As VbMsgBoxResult
    Dim organizerName As String
    Dim organizerCount As Integer
    Dim organizerProcessed As Boolean
    
    ' Initialize Outlook objects
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' Check if an item is selected in the calendar
    If Outlook.Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please highlight a meeting in your Outlook calendar first.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the selected item is an appointment
    If TypeName(Outlook.Application.ActiveExplorer.Selection.Item(1)) <> "AppointmentItem" Then
        MsgBox "Please highlight a meeting in your Outlook calendar first.", vbExclamation
        Exit Sub
    End If
    
    ' Get the selected appointment
    Set olAppointment = Outlook.Application.ActiveExplorer.Selection.Item(1)
    
    ' Get the organizer's name
    organizerName = olAppointment.GetOrganizer.Name
    organizerCount = 0
    organizerProcessed = False
    
    ' Count the number of times the organizer appears in the recipient list
    For Each olRecipient In olAppointment.Recipients
        If olRecipient.Name = organizerName Then
            organizerCount = organizerCount + 1
        End If
    Next olRecipient
    
    ' Ask the user if they want to include recipients with "None" response
    includeNone = MsgBox("Do you want to include recipients with 'None' response?", vbYesNo, "Include None Response")
    
    ' Ask the user if they want to include recipients with "Tentative" response
    includeTentative = MsgBox("Do you want to include recipients with 'Tentative' response?", vbYesNo, "Include Tentative Response")
    
    ' Initialize BCC list
    bccList = ""
    
    ' Loop through recipients
    For Each olRecipient In olAppointment.Recipients
        ' Include recipients based on user's choices
        If (includeNone = vbYes And olRecipient.MeetingResponseStatus = olResponseNone) Or _
           (includeTentative = vbYes And olRecipient.MeetingResponseStatus = olResponseTentative) Then
           
            ' Avoid listing the organizer twice
            If olRecipient.Name = organizerName Then
                If Not organizerProcessed Then
                    organizerProcessed = True
                Else
                    ' Skip this recipient if it's the organizer and we've already processed them
                    GoTo NextRecipient
                End If
            End If
            
            ' Add recipient to BCC list
            If bccList = "" Then
                bccList = olRecipient.Address
            Else
                bccList = bccList & ";" & olRecipient.Address
            End If
        End If
NextRecipient:
    Next olRecipient
    
    ' Check if BCC list is empty
    If bccList = "" Then
        MsgBox "No recipients with the selected responses found.", vbInformation
        Exit Sub
    End If
    
    ' Generate subject and body
    subject = "Awaiting Your Feedback on " & Chr(34) & olAppointment.Subject & Chr(34)
    emailBody = "Dear Colleague," & vbCrLf & vbCrLf & _
                "We're reaching out as a reminder regarding our previous communication. " & _
                "We've yet to receive a response or noted it as None/Tentative. " & _
                "We understand busy schedules and want to ensure our message is acknowledged. " & _
                "If this has been addressed, please disregard this message. " & _
                "We appreciate your prompt attention." & vbCrLf & vbCrLf & _
                "Best regards," & vbCrLf & _
                "Your Name" ' Change "Your Name" to your actual name or signature
    
    ' Create a new mail item
    Set olMail = olApp.CreateItem(olMailItem)
    With olMail
        .BCC = bccList
        .Subject = subject
        .Body = emailBody
        .Display ' This will display the email. Use .Send to send directly
        
        ' Check names (resolve all recipients)
        .Recipients.ResolveAll
    End With
    
    ' Cleanup
    Set olMail = Nothing
    Set olRecipient = Nothing
    Set olAppointment = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

