Operation =1
Option =0
Where ="(((tblInstalls.strInstallStatus)=\"Preparation\" Or (tblInstalls.strInstallStatu"
    "s)=\"Ready for Install\") AND ((tblInstalls.dtmInstallScheduled)<>\"\"))"
Begin InputTables
    Name ="tblInstalls"
    Name ="tblInstallEquipment"
End
Begin OutputColumns
    Expression ="tblInstalls.strCustomer"
    Expression ="tblInstalls.strAddressCity"
    Expression ="tblInstalls.strAddressState"
    Expression ="tblInstalls.strInstallStatus"
    Expression ="tblInstalls.dtmInstallScheduled"
    Expression ="tblInstalls.dtmDepartureTime"
    Expression ="tblInstalls.strDepartureFrom"
    Expression ="tblInstalls.dtmDeliveryDate"
    Expression ="tblInstallEquipment.strDescription"
    Expression ="tblInstallEquipment.strEQID"
    Expression ="tblInstallEquipment.strEquipmentType"
End
Begin Joins
    LeftTable ="tblInstalls"
    RightTable ="tblInstallEquipment"
    Expression ="tblInstalls.lngID = tblInstallEquipment.intInstall"
    Flag =1
End
Begin OrderBy
    Expression ="tblInstalls.dtmInstallScheduled"
    Flag =0
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
        dbText "Name" ="tblInstalls.strDepartureFrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmDepartureTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmInstallScheduled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strInstallStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strAddressState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strAddressCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strCustomer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.dtmDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.strDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.strEQID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.strEquipmentType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =980
    Bottom =833
    Left =-1
    Top =-1
    Right =956
    Bottom =554
    Left =0
    Top =288
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
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =370
        Top =0
        Name ="tblInstallEquipment"
        Name =""
    End
End
