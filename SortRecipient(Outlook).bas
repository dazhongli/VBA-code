
'This subroutine sort the recipient using an alphabetical order
'Dazhong Li 01/11/2016
Public Sub SortRecipients()
    With Application.ActiveInspector
        If TypeOf .CurrentItem Is Outlook.MailItem Then
            Dim olRecipients As Recipients
            Dim olRecipient As Recipient
            'we will sort the list based on the name showed in the globle address book
            Dim n_to, n_cc, n_bcc As Integer
            n_to = 0
            n_cc = 0
            n_bcc = 0
            Dim to_list, cc_list As Variant
            Set olRecipients = .CurrentItem.Recipients
            Dim strToEmails, strCcEmails, strBCcEmails As String
            For Each olRecipient In olRecipients
            If olRecipient.Type = Outlook.OlMailRecipientType.olTo Then
                to_list = Split(.CurrentItem.To, ";")
                If olRecipient.AddressEntry.Type = "EX" Then
                    strToEmails = strToEmails + to_list(n_to) & ";"
                Else
                    strToEmails = strToEmails + olRecipient.Address & ";"
                End If
                n_to = n_to + 1
            ElseIf olRecipient.Type = Outlook.OlMailRecipientType.olCC Then
                cc_list = Split(.CurrentItem.CC, ";")
                 If olRecipient.AddressEntry.Type = "EX" Then
                    strCcEmails = strCcEmails + cc_list(n_cc) & ";"
                Else 'if the name is an external name, we will use the address of the email coz name will cause trouble
                    strCcEmails = strCcEmails + olRecipient.Address & ";"
                End If
                n_cc = n_cc + 1
            ElseIf olRecipient.Type = Outlook.OlMailRecipientType.olBCC Then
                n_bcc = n_bcc + 1
                If olRecipient.AddressEntry.Type = "EX" Then
                    strBCcEmails = strBCcEmails + olRecipient.AddressEntry.GetExchangeUser.PrimarySmtpAddress & ";"
                Else
                    strBCcEmails = strBCcEmails + olRecipient.Address & ";"
                End If
            End If
        Next olRecipient

            ' Force an update if recipients have changed (DOESN'T HELP)
            .CurrentItem.Save
           'delete the last element in the array
            cc_list = Split(strCcEmails, ";")
            ReDim Preserve cc_list(UBound(cc_list) - 1)
            to_list = Split(strToEmails, ";")
            ReDim Preserve to_list(UBound(to_list) - 1)
            Set myRecipients = .CurrentItem.Recipients
            ' Create objects for To list
            Dim myRecipient, myTo, myCc As Recipient
            Dim recipientToList, recipientCcList As Object
            Set recipientToList = CreateObject("System.Collections.ArrayList")
            Set recipientCcList = CreateObject("System.Collections.ArrayList")
            ' Create new lists from To line
            Dim cc_item, to_item As Variant
            For Each cc_item In cc_list
                On Error Resume Next
                recipientCcList.Add Trim(cc_item)
            Next
            For Each to_item In to_list
                recipientToList.Add Trim(to_item)
            Next

            ' Sort the recipient lists
            recipientToList.Sort
            recipientCcList.Sort
            ' Remove all recipients so we can re-add in the correct order
            While myRecipients.Count > 0
                myRecipients.Remove 1
            Wend
            ' Create new To line
            Dim recipientName As Variant
            For Each recipientName In recipientToList
                Set myRecipient = myRecipients.Add(recipientName)
                myRecipient.Type = olTo
            Next recipientName
            For Each recipientName In recipientCcList
                Set myRecipient = myRecipients.Add(recipientName)
                myRecipient.Type = olCC
            Next recipientName
            .CurrentItem.Recipients.ResolveAll
        End If
    End With
End Sub


