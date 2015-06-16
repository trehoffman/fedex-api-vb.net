Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Net

Public Class FEDEX
    Private key As String = ConfigurationManager.AppSettings("FedExKey")
    Private pass As String = ConfigurationManager.AppSettings("FedExPassword")
    Private acc_no As String = ConfigurationManager.AppSettings("FedExAccountNumber")
    Private meter_no As String = ConfigurationManager.AppSettings("FedExMeterNumber")
    Private path As String = ConfigurationManager.AppSettings("FedExPath")
    Public xml As XmlDocument = New XmlDocument()

    'constructor
    Sub New(trackNo As String)
        xml = getFedexXmlByTrackingNumber(trackNo)
    End Sub

    'public functions
    Public Function getIdentificationNumber() As String
        Dim numString As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:TrackingNumberUniqueIdentifier")

        Return numString.Trim
    End Function

    Public Function getPieceCount() As String
        Dim numString As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageCount")

        If numString = "" Then
            numString = "0"
        End If

        Return numString.Trim
    End Function

    Public Function getPieceWeight() As String
        Dim numString As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageWeight/V6:Value")

        If numString = "" Then
            numString = "0"
        End If

        Return numString.Trim
    End Function

    Public Function getPieceDimensions() As String
        Dim length As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageDimensions/V6:Length")
        Dim width As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageDimensions/V6:Width")
        Dim height As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageDimensions/V6:Height")

        Dim dims As String = length.Trim & "x" & width.Trim & "x" & height.Trim

        Return dims
    End Function

    Public Function getShipmentWeight() As String
        Dim numString As String = getNodeValue(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:ShipmentWeight/V6:Value")

        If numString = "" Then
            numString = "0"
        End If

        Return numString.Trim
    End Function

    Public Function getShipmentPurchaseOrderNumber() As String
        Dim nodeList As XmlNodeList = getNodeList(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:OtherIdentifiers")
        Dim poStr As String = ""

        For Each n In nodeList
            If InStr(n.InnerText, "PURCHASE_ORDER") > 0 Then
                Dim namespaces As XmlNamespaceManager = New XmlNamespaceManager(xml.NameTable)
                namespaces.AddNamespace("SOAP-ENV", "http://schemas.xmlsoap.org/soap/envelope/")
                namespaces.AddNamespace("V6", "http://fedex.com/ws/track/v6")

                Dim node As XmlNode = n.SelectSingleNode("V6:Value", namespaces)
                poStr = node.InnerText
                'poStr = n.getNodeValue(n, "V6:Value")
            End If
        Next

        Return poStr.Trim
    End Function

    Public Function getShipmentReferenceNumber() As String
        Dim nodeList As XmlNodeList = getNodeList(xml, "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:OtherIdentifiers")
        Dim poStr As String = ""

        For Each n In nodeList
            If InStr(n.InnerText, "CUSTOMER_REFERENCE") > 0 Then
                Dim namespaces As XmlNamespaceManager = New XmlNamespaceManager(xml.NameTable)
                namespaces.AddNamespace("SOAP-ENV", "http://schemas.xmlsoap.org/soap/envelope/")
                namespaces.AddNamespace("V6", "http://fedex.com/ws/track/v6")

                Dim node As XmlNode = n.SelectSingleNode("V6:Value", namespaces)
                poStr = node.InnerText
                'poStr = n.getNodeValue(n, "V6:Value")
            End If
        Next

        Return poStr.Trim
    End Function

    Public Function getXmlString(trackNo As String) As String
        Dim xml_req As String = "<?xml version='1.0'?>" & _
                                        "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'>" & _
                                        "<soap:Body>" & _
                                        "<TrackRequest xmlns='http://fedex.com/ws/track/v6'>" & _
                                            "<WebAuthenticationDetail>" & _
                                                "<UserCredential>" & _
                                                    "<Key>" & key & "</Key>" & _
                                                    "<Password>" & pass & "</Password>" & _
                                                "</UserCredential>" & _
                                            "</WebAuthenticationDetail>" & _
                                            "<ClientDetail>" & _
                                                "<AccountNumber>" & acc_no & "</AccountNumber>" & _
                                                "<MeterNumber>" & meter_no & "</MeterNumber>" & _
                                            "</ClientDetail>" & _
                                            "<TransactionDetail>" & _
                                                "<CustomerTransactionId>" & "Basic_TrackRequest_q0_Internal" & "</CustomerTransactionId>" & _
                                                "<Localization>" & _
                                                    "<LanguageCode>EN</LanguageCode>" & _
                                                    "<LocaleCode>us</LocaleCode>" & _
                                                "</Localization>" &
                                            "</TransactionDetail>" & _
                                            "<Version>" & _
                                                "<ServiceId>" & "trck" & "</ServiceId>" & _
                                                "<Major>6</Major>" & _
                                                "<Intermediate>0</Intermediate>" & _
                                                "<Minor>0</Minor>" & _
                                            "</Version>" & _
                                                "<PackageIdentifier>" & _
                                                    "<Value>" & trackNo & "</Value>" & _
                                                    "<Type>TRACKING_NUMBER_OR_DOORTAG</Type>" & _
                                                "</PackageIdentifier>" & _
                                            "<IncludeDetailedScans>1</IncludeDetailedScans>" & _
                                        "</TrackRequest>" & _
                                        "</soap:Body>" & _
                                        "</soap:Envelope>"
        Dim client As New WebClient()
        client.Headers.Add("Content-Type", "application/xml")
        Dim sentByte As Byte() = System.Text.Encoding.ASCII.GetBytes(xml_req)
        Dim responseByte As Byte() = client.UploadData(path, "POST", sentByte)

        Dim responseString As String = System.Text.Encoding.ASCII.GetString(responseByte)
        Dim responseXML As New XmlDocument()
        responseXML.LoadXml(responseString)

        Return responseString
    End Function

    'private helper functions

    Private Function getFedexXmlByTrackingNumber(trackNo As String) As XmlDocument
        Dim xml_req As String = "<?xml version='1.0'?>" & _
                                        "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'>" & _
                                        "<soap:Body>" & _
                                        "<TrackRequest xmlns='http://fedex.com/ws/track/v6'>" & _
                                            "<WebAuthenticationDetail>" & _
                                                "<UserCredential>" & _
                                                    "<Key>" & key & "</Key>" & _
                                                    "<Password>" & pass & "</Password>" & _
                                                "</UserCredential>" & _
                                            "</WebAuthenticationDetail>" & _
                                            "<ClientDetail>" & _
                                                "<AccountNumber>" & acc_no & "</AccountNumber>" & _
                                                "<MeterNumber>" & meter_no & "</MeterNumber>" & _
                                            "</ClientDetail>" & _
                                            "<TransactionDetail>" & _
                                                "<CustomerTransactionId>" & "Basic_TrackRequest_q0_Internal" & "</CustomerTransactionId>" & _
                                                "<Localization>" & _
                                                    "<LanguageCode>EN</LanguageCode>" & _
                                                    "<LocaleCode>us</LocaleCode>" & _
                                                "</Localization>" &
                                            "</TransactionDetail>" & _
                                            "<Version>" & _
                                                "<ServiceId>" & "trck" & "</ServiceId>" & _
                                                "<Major>6</Major>" & _
                                                "<Intermediate>0</Intermediate>" & _
                                                "<Minor>0</Minor>" & _
                                            "</Version>" & _
                                                "<PackageIdentifier>" & _
                                                    "<Value>" & trackNo & "</Value>" & _
                                                    "<Type>TRACKING_NUMBER_OR_DOORTAG</Type>" & _
                                                "</PackageIdentifier>" & _
                                            "<IncludeDetailedScans>1</IncludeDetailedScans>" & _
                                        "</TrackRequest>" & _
                                        "</soap:Body>" & _
                                        "</soap:Envelope>"
        Dim client As New WebClient()
        client.Headers.Add("Content-Type", "application/xml")
        Dim sentByte As Byte() = System.Text.Encoding.ASCII.GetBytes(xml_req)
        Dim responseByte As Byte() = client.UploadData(path, "POST", sentByte)

        Dim responseString As String = System.Text.Encoding.ASCII.GetString(responseByte)
        Dim responseXML As New XmlDocument()
        responseXML.LoadXml(responseString)

        Return responseXML
    End Function

    Private Function getNodeList(xml As XmlDocument, nodePath As String) As XmlNodeList
        Dim namespaces As XmlNamespaceManager = New XmlNamespaceManager(xml.NameTable)
        namespaces.AddNamespace("SOAP-ENV", "http://schemas.xmlsoap.org/soap/envelope/")
        namespaces.AddNamespace("V6", "http://fedex.com/ws/track/v6")

        'Dim xPathString = "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageCount"
        'Dim oNode = xml.SelectSingleNode(xPathString, namespaces)
        'Return oNode.InnerXml.ToString

        Dim nodeList As XmlNodeList = xml.SelectNodes(nodePath, namespaces)

        Return nodeList
    End Function

    Private Function getNodeValue(xml As XmlDocument, nodePath As String) As String
        Dim namespaces As XmlNamespaceManager = New XmlNamespaceManager(xml.NameTable)
        namespaces.AddNamespace("SOAP-ENV", "http://schemas.xmlsoap.org/soap/envelope/")
        namespaces.AddNamespace("V6", "http://fedex.com/ws/track/v6")

        'Dim xPathString = "SOAP-ENV:Envelope/SOAP-ENV:Body/V6:TrackReply/V6:TrackDetails/V6:PackageCount"
        'Dim oNode = xml.SelectSingleNode(xPathString, namespaces)
        'Return oNode.InnerXml.ToString

        Dim node As XmlNode = xml.SelectSingleNode(nodePath, namespaces)

        If Not node Is Nothing Then
            Return node.InnerText.Trim
        End If

        Return ""
    End Function

    Private Function getNodeCount(xml As XmlDocument, nodePath As String) As Integer
        Dim nodeList = xml.SelectNodes(nodePath)
        Return nodeList.Count
    End Function

End Class
