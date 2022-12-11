SELECT tblInstallEquipment.strDescription, tblInstallEquipment.intOptionFor, tblInstallEquipment.ysnInStock, tblInstallEquipment.ysnReadyForInstall, tblInstallEquipment.lngID, tblInstallEquipment.intInstall, tblInstallEquipment.strEquipmentType
FROM tblInstallEquipment
WHERE (((tblInstallEquipment.strEquipmentType)="Accessory" Or (tblInstallEquipment.strEquipmentType)="Startup Supplies"));
