Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Public Class EventClassModule

    Public WithEvents oOutlook As Outlook.Application

    Private Sub oOutlook_NewMailEx(ByVal EntryIDCollection As String) Handles oOutlook.NewMailEx
        Dim objEntryIDs() As String
        Dim oItem As Outlook.MailItem
        Dim i As Integer
        Dim strEmailAddress As String
        Dim strDisplayName As String
        Dim strSubject As String
        Dim strBody As String
        Dim strCCAddress As String
        Dim iAttachmentCount As Integer

        objEntryIDs = Split(EntryIDCollection, ",")
        For i = 0 To UBound(objEntryIDs)
            oItem = oOutlook.Session.GetItemFromID(objEntryIDs(i))

            strEmailAddress = oItem.SenderEmailAddress
            strDisplayName = oItem.SenderName
            strCCAddress = oItem.CC
            strSubject = oItem.Subject
            strBody = oItem.Body
            iAttachmentCount = oItem.Attachments.Count

            Console.WriteLine("----------------------START ITEM---------------------------------")
            Console.WriteLine("Email Address : " & strEmailAddress)
            Console.WriteLine("Display Name  : " & strDisplayName)
            Console.WriteLine("CC            : " & strCCAddress)
            Console.WriteLine("Subject       : " & strSubject)
            Console.WriteLine("Body          : " & strBody)
            Console.WriteLine("Attachments   : " & iAttachmentCount.ToString)
            Console.WriteLine("-----------------------END ITEM----------------------------------")
        Next

    End Sub
End Class
