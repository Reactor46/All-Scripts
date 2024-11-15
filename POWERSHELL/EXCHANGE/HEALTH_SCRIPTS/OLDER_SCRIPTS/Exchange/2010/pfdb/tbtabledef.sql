
CREATE TABLE [dbo].[Attachments@address@domain.com] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[EntryID] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[AttachNum] [int] NULL ,
	[FileName] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ContentType] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ContentTransferEncoding] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ContentDisposition] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MessageBody] [text] COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Inbox@address@domain.com] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[EntryID] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MessageID] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[DAVhref] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[DateSent] [datetime] NULL ,
	[FromName] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[FromEmail] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ToEmail] [varchar] (800) COLLATE Latin1_General_CI_AS NULL ,
	[Subject] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[TextBodyShort] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[TextBody] [text] COLLATE Latin1_General_CI_AS NULL ,
	[hasattach] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Sent@address@domain.com] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[EntryID] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[DateSent] [datetime] NULL ,
	[ToEmail] [varchar] (800) COLLATE Latin1_General_CI_AS NULL ,
	[Subject] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[TextBody] [text] COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.[vwViewMessages@address@domain.com]
AS
SELECT     TOP 100 PERCENT dbo.[Inbox@address@domain.com].ID, dbo.[Inbox@address@domain.com].EntryID, dbo.[Inbox@address@domain.com].DAVhref, 
                      dbo.[Inbox@address@domain.com].MessageID, dbo.[Inbox@address@domain.com].FromName, dbo.[Inbox@address@domain.com].DateSent, 
                      dbo.[Inbox@address@domain.com].FromEmail, dbo.[Inbox@address@domain.com].ToEmail, dbo.[Inbox@address@domain.com].Subject, 
                      dbo.[Inbox@address@domain.com].TextBodyShort, dbo.[Inbox@address@domain.com].TextBody, dbo.[Inbox@address@domain.com].hasattach, 
                      dbo.[Attachments@address@domain.com].AttachNum, dbo.[Attachments@address@domain.com].FileName
FROM         dbo.[Attachments@address@domain.com] RIGHT OUTER JOIN
                      dbo.[Inbox@address@domain.com] ON dbo.[Attachments@address@domain.com].EntryID = dbo.[Inbox@address@domain.com].EntryID
ORDER BY dbo.[Inbox@address@domain.com].DateSent DESC




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

