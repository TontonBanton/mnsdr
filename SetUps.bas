Attribute VB_Name = "SetUps"
Public Sub SetlvwDR()
With FormDR.lvwDR
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", .Width * 0.05
        .ColumnHeaders.Add , , "HDate", .Width * 0.1
        .ColumnHeaders.Add , , "Area", .Width * 0.12
        .ColumnHeaders.Add , , "Specie", .Width * 0.2
        .ColumnHeaders.Add , , "Hill#", .Width * 0.08
        .ColumnHeaders.Add , , "Bolt#", .Width * 0.08
        .ColumnHeaders.Add , , "Size", .Width * 0.12
        .ColumnHeaders.Add , , "Pcs", .Width * 0.12
        .ColumnHeaders.Add , , "Bd.Ft.", .Width * 0.12
        .ColumnHeaders.Item(9).Alignment = lvwColumnRight
End With
End Sub
Public Sub SetInventory()
With FormInv.lvwInv
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Date Cut", .Width * 0.13
    .ColumnHeaders.Add , , "Area", .Width * 0.13
    .ColumnHeaders.Add , , "Specie", .Width * 0.23
    .ColumnHeaders.Add , , "Hill", .Width * 0.08
    .ColumnHeaders.Add , , "Bolt", .Width * 0.08
    .ColumnHeaders.Add , , "Size", .Width * 0.13
    .ColumnHeaders.Add , , "Pcs", .Width * 0.07
    .ColumnHeaders.Add , , "Bd. Ft.", .Width * 0.13
    .ColumnHeaders.Add , , "#", .Width * 0
End With
End Sub
Public Sub SetDRInventory()
With FormDR.lvwInv
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Date Cut", .Width * 0.13
    .ColumnHeaders.Add , , "Area", .Width * 0.13
    .ColumnHeaders.Add , , "Specie", .Width * 0.23
    .ColumnHeaders.Add , , "Hill", .Width * 0.08
    .ColumnHeaders.Add , , "Bolt", .Width * 0.08
    .ColumnHeaders.Add , , "Size", .Width * 0.13
    .ColumnHeaders.Add , , "Pcs", .Width * 0.07
    .ColumnHeaders.Add , , "Bd. Ft.", .Width * 0.13
    .ColumnHeaders.Add , , "#", .Width * 0
End With
    FormDR.FrameInv.Visible = True
End Sub
Public Sub SetSearch()
With FormDR.FrameDRSearch
    .Top = FormDR.lvwDR.Top: .Left = FormDR.lvwDR.Left: .Width = FormDR.lvwDR.Width: .Height = FormDR.lvwDR.Height: .Visible = True
End With
With FormDR.lvwDRSearch
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Number", .Width * 0.12
    .ColumnHeaders.Add , , "Date", .Width * 0.08
    .ColumnHeaders.Add , , "Buyer", .Width * 0.2
    .ColumnHeaders.Add , , "Bd. Ft.", .Width * 0.12
    .ColumnHeaders.Add , , "Cu. Mt.", .Width * 0.18
    .ColumnHeaders.Add , , "Remarks", .Width * 0.23
    '.ColumnHeaders.Item().Alignment = lvwColumnRight
End With
End Sub
Public Sub SetDelivered()
With FormDR.frameDelivered
    .Visible = True: .Top = FormDR.Frame1.Top: .Left = 14850: .Width = 10300: .Height = 7900
End With
    With FormDR.lvwDelivered
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "Customer", .Width * 0.95
     End With
End Sub
Public Sub SetDestination()
With FormDR.FrameDestination
    .Visible = True: .Top = FormDR.Frame1.Top: .Left = 14850: .Width = 10300: .Height = 7900
End With
    With FormDR.lvwDestination
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "", .Width * 0.9
    End With
    With FormDR.lvwVehicle
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "", .Width * 0.58
         .ColumnHeaders.Add , , "", .Width * 0.38
     End With
End Sub
Public Sub ClearDRDetails()
With FormDR
    .txtDRProduct.Text = ""
    .txtDRProductSearch.Text = ""
    .txtDRArea.Text = ""
End With
End Sub
Public Sub ClearBox()
With FormDR
    .lvwDR.ListItems.Clear
    .txtDRNum.Text = "": .txtDRDate.Text = "__/__/____"
    .txtDRDelivered.Text = "": .txtDeliverSearch.Text = ""
    .txtDRDestination.Text = "": .txtDRDriver.Text = ""
    .txtVehicleSearch.Text = "": .txtDestinationSearch.Text = ""
    .txtDRRemarks.Text = ""
    .txtDRTotalPcs.Text = ""
    .txtDRTotalBdFt.Text = ""
    .txtDRTotalCuM.Text = ""
End With
End Sub
Public Sub ClearFrame()
With FormDR
    .FrameInv.Visible = False
    .FrameDRDetails.Visible = False: .FrameDRSearch.Visible = False
    .frameDelivered.Visible = False: .FrameProduct.Visible = False: .FrameDestination.Visible = False
End With
End Sub
