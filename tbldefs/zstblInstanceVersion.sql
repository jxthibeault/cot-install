CREATE TABLE [zstblInstanceVersion] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [strVersion] VARCHAR (255),
  [strDatePublished] VARCHAR (255),
  [strChangelog] LONGTEXT 
)
