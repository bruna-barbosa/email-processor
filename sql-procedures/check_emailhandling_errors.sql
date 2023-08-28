USE [MasterDB]
GO

/****** Object:  StoredProcedure [dbo].[Check_EmailHandling_Errors]    Script Date: 28/08/2023 12:13:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		Bruna Duarte
-- Create date: 2023-03-08
-- Description:	This procedure sends an email if there are any errors
-- present in the emails with content sent to the server 
-- =============================================
  
CREATE PROCEDURE [dbo].[Check_EmailHandling_Errors] 
	-- Add the parameters for the stored procedure here	
	@EmailSender NVARCHAR(255),
	@SubjectSender NVARCHAR(255),
	@mailbody_content nvarchar(MAX)
	
AS
BEGIN
	
	BEGIN

		DECLARE @mailsubject AS NVARCHAR(255);
		DECLARE @NotificationID INT;
		DECLARE @EmailRecipientsTO NVARCHAR(MAX);
		DECLARE @EmailRecipientsCC NVARCHAR(MAX);
		DECLARE @mailbody NVARCHAR(MAX);		
		DECLARE @replysubject NVARCHAR(MAX);
		
		--  set NotificationID 4 Email Handling Error
		SET @NotificationID = 29;

    -- Emailing to:
		SET @EmailRecipientsTO = @EmailSender;
    -- Email CC'd to:
		SET @EmailRecipientsCC = MasterDB.dbo.GetEmailRecipientsCC4NotificationID(@NotificationID);
		
		-- 2023-05-04 Bruna: mailbody_content variable obtained from the .py script is built dinamically
		
		SET @mailbody = @mailbody_content;
		
		SET @replysubject = 'RE: ' + @SubjectSender;
		
		EXEC msdb.dbo.sp_send_dbmail 
			@profile_name = 'MailProfile', 
			@recipients = @EmailRecipientsTO, 
			@copy_recipients = @EmailRecipientsCC,
			@subject = @replysubject, 
			@body = @mailbody, 
			@body_format = 'HTML' 
		
		Print 'Sent email'
	END





END
GO


