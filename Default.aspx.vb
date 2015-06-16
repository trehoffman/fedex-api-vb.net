
Partial Class _Default
    Inherits System.Web.UI.Page

    Sub GetFedExTrackingDetails()
        Dim trackNo As String = fedexTrackNo.Value
        Dim fedex As New FEDEX(trackNo)

        'Response.Write(fedex.getXmlString(trackNo))
        'Response.Write(fedex.getAllTrackingNumbersInShipment(trackNo))

        If fedex.getIdentificationNumber() <> "" Then
            shipID.InnerText = fedex.getIdentificationNumber()
            shipPieces.InnerText = fedex.getPieceCount()
            pieceWeight.InnerText = fedex.getPieceWeight()
            trackingNumber.InnerText = trackNo
            pieceDimensions.InnerText = fedex.getPieceDimensions()
            shipWeight.InnerText = fedex.getShipmentWeight()
            PONumber.InnerText = fedex.getShipmentPurchaseOrderNumber()
            ReferenceNumber.InnerText = fedex.getShipmentReferenceNumber()
            MessageArea.Style("display") = "none"
            pieceDimensionsArea.Style("display") = "block"
            PONumberArea.Style("display") = "block"
            ReferenceNumberArea.Style("display") = "block"
            ShowResults.Style("display") = "block"
        Else
            MessageArea.InnerText = "Tracking Number " & trackNo & " not found."
            MessageArea.Style("display") = "block"
            ShowResults.Style("display") = "none"
        End If
    End Sub
End Class
