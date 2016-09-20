$FileName = "C:\CSV2\largest.csv"
if (Test-Path $FileName) {
  Remove-Item $FileName
}

$server = "GOF-SRV-PROD-R"
$database = "ReportServer"
$query = " SELECT TOP 10
        REPLACE(        
            REPLACE(    
                REPLACE(    
                    REPLACE(
                        REPLACE(
                            REPLACE(ReportPath,'/Reportings/rpt',''),
                        '_Global_Free_Full',''),
                    'Global_Free',''),
                'Global','') ,
            '_',''),
        'FewColumns','') AS ReportType
        ,CASE [Format] 
            WHEN 'RPL' THEN 'Preview'
            WHEN 'Excel' THEN 'Download' END AS Activity,
        REPLACE(    
            REPLACE(    
                REPLACE(    
                    REPLACE(
                        REPLACE(convert(varchar(max),[Parameters]),'&',' || '),
                    '%','-'),
                '2F',''),
            '-2000-3A00-3A00',''),
        '2C','') AS DownloadDetails
        ,TimeStart
        ,TimeEnd
        ,convert(decimal(5,2),
            (convert(float,TimeDataRetrieval)+convert(float,TimeProcessing)+convert(float,TimeRendering))
        /60/60) AS TimeInSecs
        ,ByteCount/1000 AS SizeInKB
    FROM ExecutionLog2
WHERE TimeStart >= GETUTCDATE() - 1 AND TimeStart < GETUTCDATE() 
AND ReportPath LIKE '%/Reportings/%'
ORDER BY SizeInKB DESC"


if ($args.length -gt 0)
{
    $query = $args[0]
}

# Update this with the actual path where you want data dumped
$extractFile = @"
C:\CSV2\largest.csv
"@

# If you have to use users and passwords, my condolences
$connectionTemplate = "Data Source={0};Integrated Security=SSPI;Initial Catalog={1};"
$connectionString = [string]::Format($connectionTemplate, $server, $database)
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

$command = New-Object System.Data.SqlClient.SqlCommand
$command.CommandText = $query
$command.Connection = $connection

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $command
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$connection.Close()

# dump the data to a csv
$DataSet.Tables[0] | Export-Csv $extractFile

Write-host "Running SQL Query to export to ANSI CSV file"

Invoke-Sqlcmd -query $dbQuery -ServerInstance $instanceName | export-csv $extractFile


$dbMailQuery3 = "execute msdb..sp_send_dbmail

 @profile_name = 'OperatorEmail', 

 @recipients = 'sysadmin@gofluent.com;dba@gofluent.com;fsamson@gofluent.com;jvillegas@gofluent.com', 

 @subject = 'Largest download time-GOF-SRV-PROD-R', 

 @body_format = 'TEXT', 

 @body = '--LARGEST DOWNLOAD--', 

 @file_attachments = 'C:\CSV2\largest.csv' "


Write-host "Sending email using DBMAIL, including .CSV file as attachment"

Invoke-Sqlcmd -query $dbMailQuery3 -ServerInstance $instanceName 

Write-host "After sending,the CSV File has been move to C:\CSV2\Moveitem"

Move-Item 'C:\CSV2\largest.csv' -Destination 'C:\CSV2\Moveitem' -Force