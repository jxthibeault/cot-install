SELECT tblInstalls.strCustomer, tblInstalls.strAddressCity, tblInstalls.strAddressState, tblInstalls.strInstallStatus, tblInstalls.dtmInstallScheduled, tblInstalls.dtmDepartureTime, tblInstalls.strDepartureFrom
FROM tblInstalls
WHERE (((tblInstalls.strInstallStatus)="Preparation" Or (tblInstalls.strInstallStatus)="Ready for Install") AND (Not (tblInstalls.dtmInstallScheduled)=""))
ORDER BY tblInstalls.dtmInstallScheduled;
