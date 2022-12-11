Operation =1
Option =0
Where ="(((tblInstallEquipment.strEquipmentType)=\"Accessory\" Or (tblInstallEquipment.s"
    "trEquipmentType)=\"Startup Supplies\"))"
Begin InputTables
    Name ="tblInstallEquipment"
End
Begin OutputColumns
    Expression ="tblInstallEquipment.strDescription"
    Expression ="tblInstallEquipment.intOptionFor"
    Expression ="tblInstallEquipment.ysnInStock"
    Expression ="tblInstallEquipment.ysnReadyForInstall"
    Expression ="tblInstallEquipment.lngID"
    Expression ="tblInstallEquipment.intInstall"
    Expression ="tblInstallEquipment.strEquipmentType"
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
        dbText "Name" ="tblInstallEquipment.[strDescription]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.strDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.ysnInStock"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.ysnReadyForInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.intOptionFor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.lngID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstallEquipment.intInstall"
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
    Right =964
    Bottom =257
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblInstallEquipment"
        Name =""
    End
End
