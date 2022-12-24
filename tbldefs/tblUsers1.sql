CREATE TABLE [tblUsers1] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [strUsername] VARCHAR (255) CONSTRAINT [strUsername] UNIQUE,
  [strDisplayName] VARCHAR (255),
  [strPassword] VARCHAR (255),
  [strAccountType] VARCHAR (255),
  [strTitle] VARCHAR (255)
)
