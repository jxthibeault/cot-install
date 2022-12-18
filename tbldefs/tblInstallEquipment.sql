CREATE TABLE [tblInstallEquipment] (
  [lngID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [intQuantity] SHORT ,
  [strDescription] VARCHAR (255),
  [strSerialNumber] VARCHAR (255),
  [strEQID] VARCHAR (255),
  [intMeterMono] VARCHAR (255),
  [intMeterColor] VARCHAR (255),
  [intOptionFor] VARCHAR (50),
  [ysnInStock] BIT ,
  [ysnReadyForInstall] BIT ,
  [strEquipmentType] VARCHAR (255),
  [intInstall] LONG ,
  [strLocation] VARCHAR (255),
  [strIpAddress] VARCHAR (255)
)
