Imports Microsoft.Office.Interop.Outlook

Public Class ThisAddIn

    Public itm As Outlook.MailItem
    Dim sentItems As Outlook.Items
    Dim sentFolder As Outlook.Folder

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ASL_Tools.app = Me.Application

        'get the asl store.
        'if there is a store then keep the reference in the asl_tools
        'if there is a store then get the sent items folder as areference
        'register the sent item as a trigger event
        'if there isnt one then disable the form
        ASL_Tools.Get_ASL_Store()
        If Not (IsNothing(ASL_Tools.aslStore)) Then

            'get the sentmail folder reference in the asl store
            Dim sntMailFolder As Outlook.Folder = ASL_Tools.Get_ASL_Store_SentItemsFolder()

            'if it is found then continue further otherwise disable everthing
            If Not (IsNothing(sntMailFolder)) Then

                'if the sent items folder is found then create an event
                'when an item is added to the folder it will store the email
                sentItems = sntMailFolder.Items
                AddHandler sentItems.ItemAdd, AddressOf Application_ItemSendAdd

                'check to see if the discipline is set yet. if it is enable all forms.
                Dim tmpdisc As String = ASL_Tools.get_discipline()
                If Not (tmpdisc = "") Then
                    ASL_Tools.enable_discipline(tmpdisc)
                Else
                    ASL_Tools.disable_discipline()
                End If
            Else

                ASL_Tools.disable_discipline()
            End If
        Else
            ASL_Tools.disable_discipline()
        End If

    End Sub

    Private Sub Application_ItemSendAdd(ByVal NewEmailItem As Object)
        'when the message is added to the sent items folder.
        Dim sentMessageItem As Outlook.MailItem = CType(NewEmailItem, Outlook.MailItem)
        If Not sentMessageItem Is Nothing Then
            'check each item in the sent items and process
            'if the mail property processed is false then it needs to be moved to a project directory
            'if the mail property stored is false then it needs to be moved to the server.
            ASL_Tools.ProcessSentMessages()
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_ItemSend(Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend
        If ASL_Tools.aslDiscipline = "" Then
            MsgBox("No Discipline set." & vbLf & "Click the Change Discipline button on the ASL Ribbon bar and set the Discipline so you can record emails", vbCritical, "Error")
            Exit Sub
        End If

        'get the senders domain name.
        'if the domain name is anthony-seaman then enable the store functionality.
        Dim senderDomain As String = ASL_Tools.Get_Domain_From_Address(Item.SenderEmailAddress)
        If senderDomain = "asltd.com" Then
            Dim frm As form_EmailSend = New form_EmailSend
            frm.button_record.Enabled = False
            frm.ShowDialog()

            If Not (frm.proj = "") Then
                'the user selected a project and wants to store it.
                'set the message properties

                Dim ma As class_MailItemTools = New class_MailItemTools
                ma.maItem = Item
                ma.proj = frm.proj
                ma.Set_PropertyAccessorObj()

                'Dim msgProp As class_MailProperties = New class_MailProperties
                'msgProp.Set_Category("Project=" & frm.proj, Item)
                'msgProp.Set_ProjectProperty(frm.proj, Item)
                'msgProp.Set_MessageTypeProperty("se", Item)
                'msgProp.Set_ProcessedProperty("False", Item)
                'msgProp.Set_StoredProperty("False", Item)
                'msgProp.Set_PropertyAccessorObj(Item)
            End If

            frm.Close()
        End If
    End Sub

    Public Sub Application_ItemRecord()
        If Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        'get the first selected item.
        'if the store the item resides in is in the domain name asltd.com continue

        Dim emList As List(Of Outlook.MailItem) = New List(Of Outlook.MailItem)
        For Each it As Outlook.MailItem In Application.ActiveExplorer.Selection
            emList.Add(it)
        Next

        Dim lastFld As String = emList.Item(0).Parent.fullfolderpath
        lastFld = lastFld.Substring(lastFld.Length - 5, 5)

        If Not (lastFld.ToUpper = "INBOX") Then Exit Sub


        Dim senderDomain As String = ASL_Tools.Get_Domain_From_Address(emList.Item(0).Parent.store.displayname.ToString)

        If senderDomain = "asltd.com" Then
            Dim frm As form_EmailRecord = New form_EmailRecord
            frm.button_record.Enabled = False
            frm.ShowDialog()

            If Not (frm.proj = "") Then
                'the user selected a project and wants to store it.
                'set the message properties

                For Each itm As Outlook.MailItem In emList
                    If (TypeOf itm Is Outlook.MailItem) Then
                        Dim ma As class_MailItemTools = New class_MailItemTools
                        ma.maItem = itm

                        ASL_Tools.SaveMessage(ma, frm.proj)

                    End If
                Next

            End If

            frm.Close()
        End If

    End Sub

End Class
