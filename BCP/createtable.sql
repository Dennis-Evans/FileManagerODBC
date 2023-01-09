CREATE TABLE [schemaname].[tablename](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[idNumber] [int] NOT NULL,
	[realNumber] [float] NOT NULL,
	[fixstr] [char](30) NOT NULL,
	[varStr] [varchar](500) NOT NULL,
	[DateValue] [date] NOT NULL,
	[bitValue] [bit] NOT NULL,
	[tinyvalue] [tinyint] NOT NULL,
	[realvalue] [real] NOT NULL,
	[smallint] [smallint] NOT NULL,
	[datet] [datetime] NULL,
	[timevalie] [time](3) NOT NULL,
 CONSTRAINT [PK_insertTable] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
