Operation =1
Option =0
Where ="(((tblInstalls.strInstallStatus)=\"Preparation\" Or (tblInstalls.strInstallStatu"
    "s)=\"Ready for Install\" Or (tblInstalls.strInstallStatus)=\"Ready for Install\""
    "))"
Begin InputTables
    Name ="tblInstallEquipment"
    Name ="tblInstalls"
End
Begin OutputColumns
    Expression ="tblInstallEquipment.[intQuantity]"
    Expression ="tblInstallEquipment.[strDescription]"
    Expression ="tblInstallEquipment.[strSerialNumber]"
    Expression ="tblInstallEquipment.[strEquipmentType]"
    Expression ="tblInstallEquipment.[intInstall]"
    Expression ="tblInstalls.strInstallStatus"
End
Begin Joins
    LeftTable ="tblInstalls"
    RightTable ="tblInstallEquipment"
    Expression ="tblInstalls.lngID = tblInstallEquipment.intInstall"
    Flag =1
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
        dbText "Name" ="tblInstallEquipment.[intQuantity]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInstalls.strInstallStatus"
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
    Right =1234
    Bottom =588
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =306
        Top =0
        Name ="tblInstallEquipment"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblInstalls"
        Name =""
    End
End
