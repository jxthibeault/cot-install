SELECT tblInstallEquipment.lngID, tblInstallEquipment.strDescription, tblInstallEquipment.strSerialNumber, tblInstallEquipment.strEQID, tblInstallEquipment.intMeterMono, tblInstallEquipment.intMeterColor, tblInstallEquipment.ysnInStock, tblInstallEquipment.ysnReadyForInstall, tblInstallEquipment.intInstall, tblInstallEquipment.strEquipmentType
FROM tblInstallEquipment
WHERE (((tblInstallEquipment.strEquipmentType)="Customer Equipment"));
