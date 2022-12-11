SELECT tblInstallEquipment.intQuantity, tblInstallEquipment.strDescription, tblInstallEquipment.strSerialNumber, tblInstallEquipment.strEquipmentType, tblInstallEquipment.intInstall, tblInstalls.strInstallStatus, tblInstallEquipment.ysnInStock
FROM tblInstalls INNER JOIN tblInstallEquipment ON tblInstalls.lngID = tblInstallEquipment.intInstall
WHERE (((tblInstalls.strInstallStatus)="Preparation" Or (tblInstalls.strInstallStatus)="Ready for Install" Or (tblInstalls.strInstallStatus)="Ready for Install") AND ((tblInstallEquipment.ysnInStock)=No));
