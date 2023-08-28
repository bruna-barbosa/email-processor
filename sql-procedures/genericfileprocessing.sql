USE [MasterDB]
GO

/****** Object:  StoredProcedure [dbo].[ProcessInputFile]    Script Date: 28/08/2023 12:06:48 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[ProcessInputFile] 
	-- Add the parameters for the stored procedure here
	@InputFile NVARCHAR (MAX),
	@ProcessingDir NVARCHAR (MAX),
	@InputTableName NVARCHAR(256)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	Declare @sql nvarchar(max)

	BEGIN TRANSACTION;
	BEGIN TRY	

		Set @sql= 'DELETE FROM ['+@InputTableName+'];'
		--Print @sql
		Exec(@sql)

		Set @sql='INSERT INTO [dbo].['+@InputTableName+']
					SELECT * FROM OPENROWSET(
									''MSDASQL'',
									''Driver={Microsoft Access Text Driver (*.txt, *.csv)};DefaultDir=' +@ProcessingDir+';'',
									''SELECT * FROM "' + @InputFile + '" '')'
		--Print @sql
		Exec(@sql)

		COMMIT TRANSACTION;


		RETURN 0;	
	
	END TRY
	BEGIN CATCH
		ROLLBACK TRANSACTION;
		RETURN 1; -- Could not import the file
	END CATCH;				
		

END
GO


