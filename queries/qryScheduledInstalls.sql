SELECT tblInstalls.strCustomer, tblInstalls.strAddressCity, tblInstalls.strAddressState, tblInstalls.strInstallStatus, tblInstalls.dtmInstallScheduled, tblInstalls.dtmDepartureTime, tblInstalls.strDepartureFrom, tblInstalls.dtmDeliveryDate, tblInstallEquipment.strDescription, tblInstallEquipment.strEQID, tblInstallEquipment.strEquipmentType
FROM tblInstalls INNER JOIN tblInstallEquipment ON tblInstalls.lngID = tblInstallEquipment.intInstall
WHERE (((tblInstalls.strInstallStatus)="Preparation" Or (tblInstalls.strInstallStatus)="Ready for Install") AND ((tblInstalls.dtmInstallScheduled)<>""))
ORDER BY tblInstalls.dtmInstallScheduled;
