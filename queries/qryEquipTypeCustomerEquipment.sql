SELECT tblInstallEquipment.lngID, tblInstallEquipment.strDescription, tblInstallEquipment.strSerialNumber, tblInstallEquipment.strEQID, tblInstallEquipment.intMeterMono, tblInstallEquipment.intMeterColor, tblInstallEquipment.ysnInStock, tblInstallEquipment.ysnReadyForInstall, tblInstallEquipment.intInstall, tblInstallEquipment.strEquipmentType, tblInstallEquipment.strLocation, tblInstallEquipment.strIpAddress
FROM tblInstallEquipment
WHERE (((tblInstallEquipment.strEquipmentType)="Customer Equipment"));
