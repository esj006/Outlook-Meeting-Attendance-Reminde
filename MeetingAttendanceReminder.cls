VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub SendEmailToSelectedResponders()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
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
    subject = "Awaiting Your Feedback on " & Chr(34) & olAppointment.subject & Chr(34)
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
        .subject = subject
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
    Dim olResponseStatus As olResponseStatus
    Dim tempRecipient As Recipient
    
    ' Create a temporary recipient to get the response status
    Set tempRecipient = appointment.Recipients.Add(oRecipient.Address)
    tempRecipient.Resolve
    olResponseStatus = tempRecipient.MeetingResponseStatus
    appointment.Recipients.Remove appointment.Recipients.Count ' Remove temporary recipient
    
    ' Include recipients based on user's choices
    If (includeNone = vbYes And olResponseStatus = olResponseNone) Or _
       (includeTentative = vbYes And olResponseStatus = olResponseTentative) Then
       
        ' Avoid listing the organizer twice
        If oRecipient.Name = organizerName Then
            If Not organizerProcessed Then
                organizerProcessed = True
            Else
                ' Skip this recipient if it's the organizer and we've already processed them
                Exit Sub
            End If
        End If
        
        ' Add recipient to BCC list
        If bccList = "" Then
            bccList = oRecipient.Address
        Else
            bccList = bccList & ";" & oRecipient.Address
        End If
    End If
End Sub


