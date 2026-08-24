Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Public Class Form1

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Register_Event_Handler()
    End Sub

    Public Sub Register_Event_Handler()
        Dim oEvents As New EventClassModule

        oEvents.oOutlook = New Outlook.Application
    End Sub
End Class
