SELECT tblInstallEquipment.lngID, tblInstallEquipment.strDescription, tblInstallEquipment.intQuantity, tblInstallEquipment.ysnInStock, tblInstallEquipment.intInstall, tblInstallEquipment.strEquipmentType
FROM tblInstallEquipment
WHERE (((tblInstallEquipment.strEquipmentType)="Customer Spare Supplies")) OR (((tblInstallEquipment.strEquipmentType)="Technician Equipment"));
