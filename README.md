# Outlook Meeting Attendance Reminder

This repository contains a VBA macro for Microsoft Outlook that helps send reminder emails to meeting attendees based on their response status. The macro can be used to send emails to attendees who have not responded ("None") or who have tentatively accepted the meeting ("Tentative").

## Features
- Automatically generate and send reminder emails to specific attendees.
- Customize email subject and body.
- Avoid sending duplicate emails to the meeting organizer.
- Easy integration with Outlook Ribbon for quick access.

## Installation

### Download the VBA code:
1. Download the `MeetingAttendanceReminder.cls` file from the repository.

### Open Outlook Visual Basic for Applications (VBA) Editor:
1. Open Outlook.
2. Press `Alt + F11` to open the VBA Editor.

### Import the VBA code into ThisOutlookSession:
1. In the Project Explorer, find and double-click on `ThisOutlookSession` under `Project1 (VbaProject.otm)`.
2. Copy the code from `MeetingAttendanceReminder.cls` and paste it into the `ThisOutlookSession` module.
3. Alternatively, you can directly copy the provided code from the repository's README and paste it into `ThisOutlookSession`.

### Customize the Outlook Ribbon:
1. Download the `MeetingAttendanceReminderRibbon.exportedUI` file from the repository.
2. Go to `File > Options > Customize Ribbon`.
3. Click on "Import/Export" and select "Import customization file".
4. Choose the `MeetingAttendanceReminderRibbon.exportedUI` file to add a button in the Outlook Ribbon for easy access to the macro.

## Usage
1. Highlight a meeting in your Outlook calendar.
2. Click on the "Send Reminder Email" button in the custom group on the Ribbon.
3. Choose whether to include attendees with "None" or "Tentative" response statuses.
4. The macro will generate and display an email with the appropriate recipients in the BCC field.

## Files
- `MeetingAttendanceReminder.cls`: The VBA macro code.
- `MeetingAttendanceReminderRibbon.exportedUI`: The exported UI customization file for the Outlook Ribbon.

## Version
- Current version: 0.1-beta

## Version History
### 0.1-beta
- Initial release.

### 0.2-beta
- **Handling Distribution Lists:** The macro now correctly expands distribution lists and retrieves the response status of each individual member, ensuring accurate tracking.
- **Resolve Recipients:** Added functionality to resolve all recipients' addresses, removing any unresolved or invalid addresses automatically to prevent errors.
- **Enhanced Response Status Checking:** The macro creates a temporary recipient to accurately fetch the response status for each member of a distribution list, then removes the temporary recipient to maintain a clean state.

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
    
    ' Ask the user if they want to include recipients with "None" response
    includeNone = MsgBox("Do you want to include recipients with 'None' response?", vbYesNo, "Include None Response")
    
    ' Ask the user if they want to include recipients with "Tentative" response
    includeTentative = MsgBox("Do you want to include recipients with 'Tentative' response?", vbYesNo, "Include Tentative Response")
    
    ' Initialize BCC list
    bccList = ""
    
    ' Loop through recipients
    For Each olRecipient In olAppointment.Recipients
        ' Expand distribution lists
        If olRecipient.AddressEntry.GetExchangeUser Is Nothing Then
            Dim oMembers As AddressEntries
            Dim oMember As AddressEntry
            
            Set oMembers = olRecipient.AddressEntry.Members
            If Not oMembers Is Nothing Then
                For Each oMember In oMembers
                    AddRecipientToBCC olAppointment, oMember, bccList, includeNone, includeTentative, organizerName, organizerProcessed
                Next oMember
            End If
        Else
            AddRecipientToBCC olAppointment, olRecipient.AddressEntry, bccList, includeNone, includeTentative, organizerName, organizerProcessed
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
        Dim rcp As Recipient
        For Each rcp In .Recipients
            If Not rcp.Resolved Then
                MsgBox "Could not resolve recipient: " & rcp.Name, vbExclamation
                rcp.Delete
            End If
        Next rcp
    End With
    
    ' Cleanup
    Set olMail = Nothing
    Set olRecipient = Nothing
    Set olAppointment = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

Sub AddRecipientToBCC(appointment As AppointmentItem, oRecipient As AddressEntry, ByRef bccList As String, includeNone As VbMsgBoxResult, includeTentative As VbMsgBoxResult, organizerName As String, ByRef organizerProcessed As Boolean)
    Dim olRecipient As Outlook.Recipient
    Dim olResponseStatus As OlResponseStatus
    Dim tempRecipient As Recipient
    
    ' Create a temporary recipient to get the response status
    Set tempRecipient = appointment.Recipients.Add(oRecipient.Address)
    tempRecipient.Resolve
    olResponseStatus = tempRecipient.MeetingResponseStatus
    appointment.Recipients.Remove appointment.Recipients.Count ' Remove temporary recipient
    
    ' Include recipients based on user's choices
    If (includeNone = vbYes And olResponseStatus = olResponseNone) Or (includeTentative = vbYes And olResponseStatus = olResponseTentative) Then
        If oRecipient.Name = organizerName Then
            If Not organizerProcessed Then
                organizerProcessed = True
            Else
                Exit Sub
            End If
        End If
        
        If bccList = "" Then
            bccList = oRecipient.Address
        Else
            bccList = bccList & ";" & oRecipient.Address
        End If
    End If
End Sub
