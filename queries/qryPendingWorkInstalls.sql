﻿SELECT tblInstalls.lngID, tblInstalls.strCustomer, tblInstalls.strInstallStatus, tblInstalls.strAddressStreet, tblInstalls.strAddressCity, tblInstalls.strAddressState, tblInstalls.strAddressZIP, tblInstalls.strSalesRep, tblInstalls.dtmDateReceived, tblInstalls.strContactName, tblInstalls.strContactPhone, tblInstalls.strContactEmail, tblInstalls.strITContactName, tblInstalls.strITContactPhone, tblInstalls.strITContactEmail, tblInstalls.memDeliveryNotes, tblInstalls.ysnFMAuditCreated, tblInstalls.ysnStairsRequired, tblInstalls.ysnElevatorRequired, tblInstalls.ysnDockDelivery, tblInstalls.memDeploymentInfo, tblInstalls.memNetworkInfo, tblInstalls.memSpecialInstructions, tblInstalls.memNotes, tblInstalls.dtmInstallScheduled, tblInstalls.dtmDepartureTime, tblInstalls.strDepartureFrom, tblInstalls.memPostInstallNotes, tblInstalls.dtmDeliveryDate, tblInstalls.strDeliveryMethod
FROM tblInstalls
WHERE (((tblInstalls.strInstallStatus)="Preparation"));
