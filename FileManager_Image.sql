declare 
@String Varchar(max), 
@RootPath VARCHAR(255),
@Filename VARCHAR(100),
@objFileSystem int,
@objTextStream int,
@objErrorObject int,
@strErrorMessage Varchar(1000),
@Command varchar(1000),
@hr int,
@fileAndPath varchar(800),
@subDirectory varchar(200),
@FullPath varchar(2000)

set @String = 'This is a fake file for QA test only!'

declare getPath cursor for select distinct fullpath
from table with (nolock)
where FullPath is not null 

order by FullPath 
open getPath 
fetch next from getPath into @FullPath
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

    --- select * from #tmp 

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

	set @Filename = 'qa1.txt'
	
	--Create table tmpImage (BinaryField image)
	
	declare @tmpImage table (BinaryField image)
	INSERT INTO @tmpImage (BinaryField) 
	SELECT BulkColumn 
	FROM Openrowset(Bulk 'C:\images\qa1.tif', Single_Blob) 
		as img
	
	
	DECLARE @SOURCEPATH VARBINARY(MAX),
	@DESTPATH VARCHAR(MAX),
	@ObjectToken INT,
	@image_ID BIGINT

	DECLARE IMGPATH CURSOR FAST_FORWARD FOR
	SELECT binaryfield from @tmpImage
	OPEN IMGPATH

	FETCH NEXT FROM IMGPATH INTO @SOURCEPATH

	WHILE @@FETCH_STATUS = 0
	BEGIN
	SET @DESTPATH = @fullpath+'\qa1.tif'

	EXEC sp_OACreate 'ADODB.Stream', @ObjectToken OUTPUT
	EXEC sp_OASetProperty @ObjectToken, 'Type', 1
	EXEC sp_OAMethod @ObjectToken, 'Open'
	EXEC sp_OAMethod @ObjectToken, 'Write', NULL, @SOURCEPATH
	EXEC sp_OAMethod @ObjectToken, 'SaveToFile', NULL, @DESTPATH, 2
	EXEC sp_OAMethod @ObjectToken, 'Close'
	EXEC sp_OADestroy @ObjectToken

	FETCH NEXT FROM IMGPATH INTO @SOURCEPATH
	END

	CLOSE IMGPATH
	DEALLOCATE IMGPATH
	
	
	
	set @FullPath = ''
	fetch next from getPath into @FullPath
end 
close getPath
deallocate getPath




