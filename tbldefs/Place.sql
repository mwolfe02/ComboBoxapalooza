CREATE TABLE [Place] (
  [PlaceID] AUTOINCREMENT CONSTRAINT [aaaPK_PlaceID] PRIMARY KEY UNIQUE NOT NULL,
  [CountryCode] VARCHAR (2),
  [ZipCode] VARCHAR (5),
  [PlaceName] VARCHAR (180),
  [StateName] VARCHAR (100),
  [StateCode] VARCHAR (2),
  [CountyName] VARCHAR (100),
  [CountyCode] VARCHAR (3),
  [Latitude] VARCHAR (10),
  [Longitude] VARCHAR (10),
  [Accuracy] VARCHAR (1),
   CONSTRAINT 
)
