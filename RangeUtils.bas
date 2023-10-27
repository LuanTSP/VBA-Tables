Attribute VB_Name = "RangeUtils"
Public Function ExpandRow(origin As range) As range
    Dim currRegionAddress As Collection, currAddress As Collection
    Set currRegionAddress = SplitAddress(origin.CurrentRegion)
    Set currAddress = SplitAddress(origin)
    
    Set ExpandRow = origin.Worksheet.range(currRegionAddress.item(1) & currAddress.item(2) & ":" & currRegionAddress.item(3) & currAddress.item(4))
End Function

Public Function ExpandRight(origin As range) As range
    Dim currRegionAddress As Collection, currAddress As Collection
    Set currRegionAddress = SplitAddress(origin.CurrentRegion)
    Set currAddress = SplitAddress(origin)
    
    Set ExpandRight = origin.Worksheet.range(currAddress.item(1) & currAddress.item(2) & ":" & currRegionAddress.item(3) & currAddress.item(4))
End Function

Public Function ExpandLeft(origin As range) As range
    Dim currRegionAddress As Collection, currAddress As Collection
    Set currRegionAddress = SplitAddress(origin.CurrentRegion)
    Set currAddress = SplitAddress(origin)
    
    Set ExpandLeft = origin.Worksheet.range(currRegionAddress.item(1) & currRegionAddress.item(2) & ":" & currAddress.item(3) & currAddress.item(4))
End Function

Public Function ExpandUp(origin As range) As range
    Dim currRegionAddress As Collection, currAddress As Collection
    Set currRegionAddress = SplitAddress(origin.CurrentRegion)
    Set currAddress = SplitAddress(origin)
    
    Set ExpandUp = origin.Worksheet.range(currAddress.item(1) & currRegionAddress.item(2) & ":" & currAddress.item(3) & currAddress.item(4))
End Function

Public Function ExpandDown(origin As range) As range
    Dim currRegionAddress As Collection, currAddress As Collection
    Set currRegionAddress = SplitAddress(origin.CurrentRegion)
    Set currAddress = SplitAddress(origin)
    
    Set ExpandDown = origin.Worksheet.range(currAddress.item(1) & currAddress.item(2) & ":" & currAddress.item(3) & currRegionAddress.item(4))
End Function

Public Function ExpandColumn(origin As range) As range
    Dim currRegionAddress As Collection, currAddress As Collection
    Set currRegionAddress = SplitAddress(origin.CurrentRegion)
    Set currAddress = SplitAddress(origin)
    
    Set ExpandColumn = origin.Worksheet.range(currAddress.item(1) & currRegionAddress.item(2) & ":" & currAddress.item(3) & currRegionAddress.item(4))
End Function

Private Function SplitAddress(rng As range) As Collection
    Dim split_data As New Collection
    
    data = Split(rng.Address, "$")
    Call split_data.Add(Trim(data(1)))
    Call split_data.Add(Trim(Split(data(2), ":")(0)))
    
    If UBound(data) = 4 Then
        Call split_data.Add(Trim(data(3)))
        Call split_data.Add(Trim(data(4)))
    Else
        Call split_data.Add(Trim(data(1)))
        Call split_data.Add(Trim(Split(data(2), ":")(0)))
    End If
    
    Set SplitAddress = split_data
End Function
