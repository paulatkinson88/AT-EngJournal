Public Class class_MailProperties
    Public proj As String
    Public timestamp As String
    Public messagetype As String
    Public processed As String
    Public stored As String

    Public Sub New()
        proj = ""
        timestamp = ""
        messagetype = ""
        processed = ""
        stored = ""
    End Sub

    Public Sub Get_AllProperties(maIt As Outlook.MailItem)
        Get_ProjectProperty(maIt)
        Get_TimeStampProperty(maIt)
        Get_MessageTypeProperty(maIt)
        Get_ProcessedProperty(maIt)
        Get_StoredProperty(maIt)
    End Sub

    Public Function Format_DateTimeStamp(maIt As Outlook.MailItem) As String
        Dim retVal As String = ""
        Dim cD As Date = maIt.ReceivedTime
        retVal = Format(cD, "yyyy-MM-dd-HHmmss")
        Return retVal
    End Function

    ''' <summary>
    ''' find the property in a mail item
    ''' if it isnt found then create it
    ''' </summary>
    ''' <param name="prop"></param>
    ''' <returns></returns>
    Private Function find_property(prop As String, maIt As Outlook.MailItem) As Boolean
        Dim retVal As Boolean = False

        For Each upro As Outlook.UserProperty In maIt.UserProperties
            If upro.Name = prop Then
                retVal = True
                Exit For
            End If
        Next

        If retVal = False Then
            maIt.UserProperties.Add(prop, Outlook.OlUserPropertyType.olText, False, Outlook.OlFormatText.olFormatTextText)
            retVal = True
        End If

        Return retVal
    End Function

    ''' <summary>
    ''' all messages use this
    ''' this is the project number associated to the message.
    ''' </summary>
    ''' <param name="st"></param>
    ''' <param name="maIt"></param>
    Public Sub Set_ProjectProperty(st As String, maIt As Outlook.MailItem)
        find_property("Project", maIt)
        proj = st
        maIt.UserProperties.Item("Project").Value = st
    End Sub

    Public Function Get_ProjectProperty(maIt As Outlook.MailItem)
        Dim retVal As String = ""
        find_property("Project", maIt)

        For Each upro As Outlook.UserProperty In maIt.UserProperties
            If upro.Name = "Project" Then
                proj = upro.Value
                retVal = proj
                Exit For
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' all messages use this
    ''' this is the receivedtime value of a message
    ''' </summary>
    ''' <param name="st"></param>
    ''' <param name="maIt"></param>
    Public Sub Set_TimeStampProperty(st As String, maIt As Outlook.MailItem)
        find_property("TimeStamp", maIt)
        timestamp = st
        maIt.UserProperties.Item("TimeStamp").Value = st
    End Sub

    Public Function Get_TimeStampProperty(maIt As Outlook.MailItem)
        Dim retVal As String = ""
        find_property("TimeStamp", maIt)

        For Each upro As Outlook.UserProperty In maIt.UserProperties
            If upro.Name = "TimeStamp" Then
                timestamp = upro.Value
                retVal = timestamp
                Exit For
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' all messages use this
    ''' shows if the message is 'se' sent or 're' received
    ''' </summary>
    ''' <param name="st"></param>
    ''' <param name="maIt"></param>
    Public Sub Set_MessageTypeProperty(st As String, maIt As Outlook.MailItem)
        find_property("MessageType", maIt)
        messagetype = st
        maIt.UserProperties.Item("MessageType").Value = st
    End Sub

    Public Function Get_MessageTypeProperty(maIt As Outlook.MailItem)
        Dim retVal As String = ""
        find_property("MessageType", maIt)

        For Each upro As Outlook.UserProperty In maIt.UserProperties
            If upro.Name = "MessageType" Then
                messagetype = upro.Value
                retVal = messagetype
                Exit For
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' Used to know if sent mail has been processed from the 
    ''' sent mail items mail box.
    ''' True or False
    ''' </summary>
    ''' <param name="st"></param>
    ''' <param name="maIt"></param>
    Public Sub Set_ProcessedProperty(st As String, maIt As Outlook.MailItem)
        find_property("Processed", maIt)
        processed = st
        maIt.UserProperties.Item("Processed").Value = st
    End Sub

    Public Function Get_ProcessedProperty(maIt As Outlook.MailItem)
        Dim retVal As String = ""
        find_property("Processed", maIt)

        For Each upro As Outlook.UserProperty In maIt.UserProperties
            If upro.Name = "Processed" Then
                processed = upro.Value
                retVal = processed
                Exit For
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' all messages use this. 
    ''' shows if the message has been stored on the server.
    ''' True or False
    ''' </summary>
    ''' <param name="st"></param>
    ''' <param name="maIt"></param>
    Public Sub Set_StoredProperty(st As String, maIt As Outlook.MailItem)
        find_property("Stored", maIt)
        stored = st
        maIt.UserProperties.Item("Stored").Value = st
    End Sub

    Public Function Get_StoredProperty(maIt As Outlook.MailItem)
        Dim retVal As String = ""
        find_property("Stored", maIt)

        For Each upro As Outlook.UserProperty In maIt.UserProperties
            If upro.Name = "Stored" Then
                stored = upro.Value
                retVal = stored
                Exit For
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' this is used by sent mail items
    ''' set the category value of the message item.
    ''' when an email is sentout set "Project=0000" as acategory
    ''' append the project number to the back of the name
    ''' retrieve this at a later date to process the message from the sent items
    ''' </summary>
    ''' <param name="st"></param>
    ''' <param name="maIt"></param>
    Public Sub Set_Category(st As String, maIt As Outlook.MailItem)
        maIt.Categories = st
    End Sub

    Public Function Get_Category(maIt As Outlook.MailItem)
        Dim retVal As String = ""
        If Not (IsNothing(maIt.Categories)) Then
            retVal = maIt.Categories
        End If

        Return retVal
    End Function

    Public Sub Set_PropertyAccessorObj(maIt As Outlook.MailItem)
        Dim prop As String = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/PropertyName"
        Dim propSt As String = "(" & proj & ")(" & timestamp & ")(" & messagetype & ")(" & processed & ")(" & stored & ")()()"
        maIt.PropertyAccessor.SetProperty(prop, propSt)
    End Sub

    Public Function Get_PropertyAccessorObj(maIt As Outlook.MailItem)
        Dim retVal As String = ""

        'PR_TRANSPORT_MESSAGE_HEADERS
        retVal = maIt.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
        Dim arObj = retVal.Split("(")
        If Not (arObj.Count = 0) Then
            proj = arObj(0)
            timestamp = arObj(1)
            messagetype = arObj(2)
            processed = arObj(3)
            stored = arObj(4)
        End If

        Return retVal
    End Function


End Class
