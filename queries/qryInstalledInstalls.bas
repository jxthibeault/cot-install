Operation =1
Option =0
Where ="(((tblInstalls.strInstallStatus)=\"Installed\"))"
Begin InputTables
    Name ="tblInstalls"
End
Begin OutputColumns
    Expression ="tblInstalls.lngID"
    Expression ="tblInstalls.strCustomer"
    Expression ="tblInstalls.strInstallStatus"
    Expression ="tblInstalls.strAddressStreet"
    Expression ="tblInstalls.strAddressCity"
    Expression ="tblInstalls.strAddressState"
    Expression ="tblInstalls.strAddressZIP"
    Expression ="tblInstalls.strSalesRep"
    Expression ="tblInstalls.dtmDateReceived"
    Expression ="tblInstalls.strContactName"
    Expression ="tblInstalls.strContactPhone"
    Expression ="tblInstalls.strContactEmail"
    Expression ="tblInstalls.strITContactName"
    Expression ="tblInstalls.strITContactPhone"
    Expression ="tblInstalls.strITContactEmail"
    Expression ="tblInstalls.memDeliveryNotes"
    Expression ="tblInstalls.ysnFMAuditCreated"
    Expression ="tblInstalls.ysnStairsRequired"
    Expression ="tblInstalls.ysnElevatorRequired"
    Expression ="tblInstalls.ysnDockDelivery"
    Expression ="tblInstalls.memDeploymentInfo"
    Expression ="tblInstalls.memNetworkInfo"
    Expression ="tblInstalls.memSpecialInstructions"
    Expression ="tblInstalls.memNotes"
    Expression ="tblInstalls.dtmInstallScheduled"
    Expression ="tblInstalls.dtmDepartureTime"
    Expression ="tblInstalls.strDepartureFrom"
    Expression ="tblInstalls.memPostInstallNotes"
    Expression ="tblInstalls.dtmDeliveryDate"
    Expression ="tblInstalls.strDeliveryMethod"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblInstalls.[lngID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.[strInstallStatus]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.[strCustomer]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.lngID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strInstallStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strCustomer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strAddressStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.memDeploymentInfo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strITContactName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strAddressCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.memNetworkInfo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strITContactPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strAddressState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.memSpecialInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strITContactEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strAddressZIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.memNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.memDeliveryNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strSalesRep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmInstallScheduled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.ysnFMAuditCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmDateReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.ysnStairsRequired"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strContactName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.ysnElevatorRequired"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strContactPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.ysnDockDelivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strContactEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmDepartureTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strDepartureFrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.memPostInstallNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strDeliveryMethod"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1250
    Bottom =833
    Left =-1
    Top =-1
    Right =1226
    Bottom =376
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblInstalls"
        Name =""
    End
End
