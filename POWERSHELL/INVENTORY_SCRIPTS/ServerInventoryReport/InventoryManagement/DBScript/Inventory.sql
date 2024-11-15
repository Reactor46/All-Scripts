USE [TEST_inventory]
GO
/****** Object:  Table [dbo].[category]    Script Date: 6/27/2014 4:18:56 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[category](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[category_name] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_category] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[dealers]    Script Date: 6/27/2014 4:18:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[dealers](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[dealer_name] [nvarchar](50) NULL,
	[dealer_address] [nvarchar](50) NULL,
 CONSTRAINT [PK_products_stock] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[products]    Script Date: 6/27/2014 4:18:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[products](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[product_name] [nvarchar](max) NOT NULL,
	[image_url] [nvarchar](max) NULL,
	[brand] [nvarchar](max) NULL,
	[category] [int] NOT NULL,
	[sub_category] [int] NOT NULL,
	[weight] [nvarchar](max) NOT NULL,
	[cost_price] [money] NOT NULL,
	[sell_price] [money] NOT NULL,
	[status] [bit] NOT NULL,
	[Stock] [int] NULL,
 CONSTRAINT [PK_products] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[selling_history]    Script Date: 6/27/2014 4:18:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[selling_history](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[dealer_id] [int] NULL,
	[product_id] [int] NOT NULL,
	[quantity] [int] NOT NULL,
	[credit] [money] NOT NULL,
	[debit] [money] NOT NULL,
	[transaction_type] [int] NOT NULL,
	[customer_info] [nvarchar](max) NULL,
	[payment_type] [nvarchar](max) NOT NULL,
	[payment_date] [datetime] NOT NULL,
	[customer_name] [nvarchar](max) NULL,
	[remarks] [nvarchar](max) NULL,
 CONSTRAINT [PK_selling_history] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[sub_category]    Script Date: 6/27/2014 4:18:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[sub_category](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[subcategory_name] [nvarchar](50) NOT NULL,
	[category] [int] NULL,
 CONSTRAINT [PK_sub_category] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[transaction_type]    Script Date: 6/27/2014 4:18:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[transaction_type](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[transactiontype] [nvarchar](max) NULL,
 CONSTRAINT [PK_transaction_type] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Users]    Script Date: 6/27/2014 4:18:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Users](
	[Id] [bigint] IDENTITY(1,1) NOT NULL,
	[UserName] [nvarchar](max) NOT NULL,
	[Password] [nvarchar](max) NOT NULL,
	[Name] [nvarchar](max) NULL,
	[Email] [nvarchar](max) NULL,
	[DOJ] [datetime] NOT NULL,
	[UpdatedOn] [datetime] NULL,
 CONSTRAINT [PK_Users] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_status]  DEFAULT ((1)) FOR [status]
GO
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_Stock]  DEFAULT ((0)) FOR [Stock]
GO
ALTER TABLE [dbo].[selling_history] ADD  CONSTRAINT [DF_selling_history_payment_date]  DEFAULT (getdate()) FOR [payment_date]
GO
ALTER TABLE [dbo].[Users] ADD  CONSTRAINT [DF_Users_DOJ]  DEFAULT (getdate()) FOR [DOJ]
GO
ALTER TABLE [dbo].[selling_history]  WITH CHECK ADD  CONSTRAINT [FK_selling_history_dealers] FOREIGN KEY([dealer_id])
REFERENCES [dbo].[dealers] ([id])
GO
ALTER TABLE [dbo].[selling_history] CHECK CONSTRAINT [FK_selling_history_dealers]
GO
