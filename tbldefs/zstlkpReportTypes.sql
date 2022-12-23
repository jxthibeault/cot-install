CREATE TABLE [zstlkpReportTypes] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [strReportTitle] VARCHAR (255),
  [strObjectName] VARCHAR (255),
  [strObjectType] VARCHAR (255)
)
