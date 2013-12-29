declare @dbObjectName varchar(50)
--set @dbObjectName = 'CaseFile_Ext_V'
set @dbObjectName = 'RepositoryData1_Ext_V'

--Go get files and put them in a tmp table
declare @tmpImage table (BinaryField image)
	INSERT INTO @tmpImage (BinaryField) 
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.tif', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.tiff', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.pdf', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.txt', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.doc', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.docx', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.xls', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.xlsx', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.zip', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.rtf', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.htm', Single_Blob) 
		as img
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\Users\ccui\TestImages\qa1.xml', Single_Blob) 
		as img

--go get the extension and put them in a temp table
print 'CHANGE DB OBJECT NAME'	
declare @tmp table (ext varchar(10))
if @dbObjectName <> 'CaseFile_Ext_V'
	insert into @tmp select distinct rtrim(ltrim(fileext)) 
		from RepositoryData1_Ext_V with (nolock)
else
	insert into @tmp values ('.tif')
	--the following sql is more accurate but too expensive
	--select distinct substring(filename, CHARINDEX('.', FileName,1), LEN(filename))
		--from CaseFile_Base with (nolock)
		
declare 
@RootPath VARCHAR(255),
@Filename VARCHAR(100),
@subDirectory varchar(200),
@FullPath varchar(2000)

print 'CHANGE DB OBJECT NAME'	
declare getPath cursor for select distinct fullpath
from RepositoryData1_Ext_V with (nolock)
where fullPath is not null order by fullPath 
open getPath 
fetch next from getPath into @fullPath
while @@FETCH_STATUS = 0 begin	
	
	Declare @pos1 int, @pos2 int, @pos3 int, @pos4 int,
	@spliton varchar(1)
	set @spliton = '\'
	
	set @pos1 = charindex( @spliton, @FullPath )
	set @pos2 = charindex( @spliton, @FullPath , @pos1+1)
	set @pos3 = charindex( @spliton, @FullPath , @pos2+1)
	set @pos4 = charindex( @spliton, @FullPath , @pos3+1)
	
	--@rootpath ends with the 4th \
	set @RootPath = SUBSTRING(@fullpath, 1, @pos4)
	
	--@subdirectory begins after the 4th slash
	set @subDirectory = SUBSTRING(@fullpath,@pos4+1, LEN(@fullpath))

	DECLARE @DirTree TABLE (subdirectory nvarchar(255), depth INT)

	--temp table to hold subdirectory names
	create table #tmp 
	(
	Id int identity(1,1),
    Data nvarchar(100)
    )
    
    delete #tmp
    
    --parse @subdirectory and put elements into temp table
    set @subDirectory = @subDirectory + '\'
    While (Charindex(@spliton,@subDirectory)>0)
    Begin
        Insert Into #tmp (data)
        Select Data = ltrim(rtrim(Substring(@subDirectory,1,Charindex(@spliton,@subDirectory)-1)))

        Set @subDirectory = Substring(@subDirectory,Charindex(@spliton,@subDirectory)+1,len(@subDirectory))
    End


	INSERT INTO @DirTree(subdirectory, depth)
	EXEC master.sys.xp_dirtree @fullPath
	
	--go thorugh temp table to create new sub direcotory if not exist
	declare @newDir sysname
	declare getTmp cursor for select data from #tmp order by id
	open getTmp
	fetch next from getTmp into @newDir
	while @@FETCH_STATUS = 0 begin
		print @newDir
		IF NOT EXISTS (SELECT 1 FROM @DirTree 
			WHERE subdirectory = @newDir)
			EXEC master.dbo.xp_create_subdir @fullPath
		
		set @newDir = ''
		fetch next from getTmp into @newDir
	end
	close getTmp
	deallocate getTmp

	drop table #tmp
			
	--create files
	declare @ext varchar(10)
	declare getExt cursor for select ext from @tmp
	open getExt
	fetch next from getExt into @ext
	while @@FETCH_STATUS = 0 begin
		set @Filename = 'qa1'+@ext
		
		declare @file VARBINARY(MAX),
		@destPath VARCHAR(MAX),
		@ObjectToken INT

		declare getFile cursor fast_forward for
		select binaryfield from @tmpImage
		open getFile

		fetch next from getFile INTO @file

		while @@FETCH_STATUS = 0
		begin
		set @destPath = @fullpath+'\'+@filename

		EXEC sp_OACreate 'ADODB.Stream', @ObjectToken OUTPUT
		EXEC sp_OASetProperty @ObjectToken, 'Type', 1
		EXEC sp_OAMethod @ObjectToken, 'Open'
		EXEC sp_OAMethod @ObjectToken, 'Write', NULL, @file
		EXEC sp_OAMethod @ObjectToken, 'SaveToFile', NULL, @destPath, 2
		EXEC sp_OAMethod @ObjectToken, 'Close'
		EXEC sp_OADestroy @ObjectToken
		
		/* delete files
		DECLARE @Result int
		DECLARE @FSO_Token int
		EXEC @Result = sp_OACreate 'Scripting.FileSystemObject', @FSO_Token OUTPUT
		EXEC @Result = sp_OAMethod @FSO_Token, 'DeleteFile', NULL, @destPath
		EXEC @Result = sp_OADestroy @FSO_Token
		*/
		
		fetch next from getFile INTO @file
		end

		close getFile
		deallocate getFile

		set @ext = ''
		fetch next from getExt into @ext
		end
	close getExt
	deallocate getExt
	
set @fullPath = ''
fetch next from getPath into @fullPath
end 
close getPath
deallocate getPath

--make sure the filename in db is in sync with what is in directories
print 'CHANGE DB OBJECT NAME'	
update RepositoryData1_BASE set FileName = 'qa1'+FileExt



