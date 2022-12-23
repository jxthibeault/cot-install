SELECT tblInstalls.strCustomer, tblInstalls.strAddressCity, tblInstalls.strAddressState, tblInstalls.strInstallStatus, tblInstalls.dtmInstallScheduled, tblInstalls.dtmDepartureTime, tblInstalls.strDepartureFrom, tblInstalls.dtmDeliveryDate
FROM tblInstalls
WHERE (((tblInstalls.strInstallStatus)="Preparation" Or (tblInstalls.strInstallStatus)="Ready for Install") AND ((tblInstalls.dtmInstallScheduled)<>""))
ORDER BY tblInstalls.dtmInstallScheduled;
