<##
#  Tax Declaration Update Tool
#  ===========================
#
#  Loads data for tax declarations from a csv file and updates records via the API.
#  Please be aware that it makes 2 API calls PER LINE of data in the csv file, so
#  may take a long time to load a large file and also slow down the Blackbaud 
#  environment. It may be a good idea to break the load into chunks or run the job
#  during a quiet period.
#
#  To use the tool:
#   1. Prepare the csv file:
#      - There must be exactly 1 column with a title that ends in "System record ID" 
#        that contains the ID (NOT the Lookup ID) of the tax declaration record to be updated.
#      - Each column to be mapped must exactly match the column ID that it corresponds to in 
#        the DataForm "Tax Declaration Edit Form 2". These IDs can be found in the DataForm 
#        metadata using Design Mode. At the time of writing, they are - 
#            | Field ID                       | Caption                  | Data type  | Descriptor                                                                |
#            |================================|==========================|============|===========================================================================|
#            | CONSTITUENTID                  |                          | Guid       |                                                                           |
#            | DECLARATIONMADE                | Made	                 | Date       |                                                                           |
#            | DECLARATIONSTARTS              | Start date               | Date       |                                                                           |
#            | DECLARATIONENDS                | End date                 | Date       |                                                                           |
#            | DECLARATIONINDICATORCODE       | Indicator                | TinyInt    | Value List                                                                |
#            | DECLARATIONSOURCECODEID        | Source                   | Guid       | Code Table (Declaration Source)                                           |
#            | CHARITYCLAIMREFERENCENUMBERID  | Reference number         | Guid       | Simple Data List (Charity Claim Reference Number By ID Simple Data List)  |
#            | SCANNEDDOCSEXIST	              | Scanned documents exist  | Boolean    |                                                                           |
#            | CONFIRMATIONSENT	              | Sent                     | Date       |                                                                           |
#            | CONFIRMATIONRETURNED	          | Returned                 | Date       |                                                                           |
#            | PAYSTAXCODE	                  | Pays tax                 | TinyInt    | Value List                                                                |
#            | TAXSTATUSCODEID	              | Status                   | Guid       | Code Table (Tax Status)                                                   |
#            | COMMENTS	                      | Comment                  | String     |                                                                           |
#      - Each column to be mapped must be valid for the data type (in the table above). Currently 
#        the tool can update Guid, Date, String, Integer and Boolean types.
#      - The tool does not map values to ids for Code Tables, Simple Data Lists or Value Lists. 
#        You will have to look these up and resolve them in the csv file before importing.
#   2. Run the tool by right-clicking the file TaxDeclarationUpdateTool.ps1 and selecting "Run with PowerShell".
#   3. Click the button "Choose file to import..." and select your csv file.
#   4. Select the Blackbaud CRM environment you want to load the data into.
#   5. Click the "Credential" button and enter your username and password for the selected environment.
#   6. If you want to, you can limit which rows are loaded by using the "Import rows from" and "to" numeric fields.
#   7. When you are ready, click the "Start import" button. Any output will be logged to the big listbox.
#      
#>


$Script:Csv = $null

$Script:Environments = @(
    @{ 'Name' = 'Staging';    'ServiceUrl' = 'https://crm79599s.sky.blackbaud.com/79599S_049d71b2-6d44-46f2-873f-76d08b145d5d/appfxwebservice.asmx'; 'Database' = '79599S'; 'Credential' = $null; }
    @{ 'Name' = 'Production'; 'ServiceUrl' = 'https://crm79599p.sky.blackbaud.com/79599P_398bba46-db35-4427-b732-b01ea09a37a9/appfxwebservice.asmx'; 'Database' = '79599P'; 'Credential' = $null; }
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '600,400'
$Form.MinimumSize                = '220,300'
$Form.Text                       = "Tax Declaration Update Tool"
$Form.TopMost                    = $false
$Form.BringToFront()
$Form.StartPosition              = 'CenterScreen'

$btnFile                         = New-Object system.Windows.Forms.Button
$btnFile.Text                    = "Choose file to import..."
$btnFile.Width                   = 280
$btnFile.Height                  = 30
$btnFile.Anchor                  = 'top,left,right'
$btnFile.Location                = New-Object System.Drawing.Point(10,25)
$btnFile.Add_Click( {
    $FileOpenForm = New-Object System.Windows.Forms.OpenFileDialog
    $FileOpenForm.Filter = "Csv files (*.csv)|*.csv|All files(*.*)|*.*"
    If( $FileOpenForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK ) {
        $Script:Csv = Import-Csv $FileOpenForm.FileName
        $btnFile.Text = $FileOpenForm.SafeFileName
        $numStartRow.Enabled = $true
        $numStartRow.Maximum = $Script:Csv.Length
        $numStartRow.Value = 1
        $numEndRow.Enabled = $true
        $numEndRow.Maximum = $Script:Csv.Length
        $numEndRow.Value = $Script:Csv.Length
    } Else {
        $Script:Csv = $null
    }
} )

$labelEnvironment                = New-Object system.Windows.Forms.Label
$labelEnvironment.Text           = "Select environment to import into"
$labelEnvironment.Width          = 190
$labelEnvironment.Height         = 20
$labelEnvironment.Anchor          = 'top,right'
$labelEnvironment.Location       = New-Object System.Drawing.Point(310,10)

$textEnvironment                 = New-Object System.Windows.Forms.ComboBox
$textEnvironment.Width           = 190
$textEnvironment.Height          = 20
$textEnvironment.Anchor          = 'top,right'
$textEnvironment.Location        = New-Object System.Drawing.Point(310,30)
$Environments | %{ $textEnvironment.Items.Add( $_.Name ) | Out-Null }
$textEnvironment.SelectedIndex   = 0

$btnLogin                        = New-Object system.Windows.Forms.Button
$btnLogin.Text                   = "Credential"
$btnLogin.Width                  = 80
$btnLogin.Height                 = 30
$btnLogin.Anchor                 = 'top,right'
$btnLogin.Location               = New-Object System.Drawing.Point(510,25)
$btnLogin.Add_Click( {
    $Script:Environments[ $textEnvironment.SelectedIndex ].Credential = Get-Credential -UserName $Username `
        -Message "Please enter your login credentials for the Blackbaud CRM $($Script:Environments[ $textEnvironment.SelectedIndex ].Name) environment"
    $Form.BringToFront()
} )

$labelStartRow                   = New-Object System.Windows.Forms.Label
$labelStartRow.Text              = "Import rows from"
$labelStartRow.Width             = 90
$labelStartRow.Height            = 20
$labelStartRow.Anchor            = 'top,left'
$labelStartRow.Location          = New-Object System.Drawing.Point(10,73)

$numStartRow                     = New-Object System.Windows.Forms.NumericUpDown
$numStartRow.Value               = $null
$numStartRow.Width               = 100
$numStartRow.Height              = 20
$numStartRow.Anchor              = 'top,left'
$numStartRow.Location            = New-Object System.Drawing.Point(100,70)
$numStartRow.Enabled             = $false

$labelEndRow                     = New-Object System.Windows.Forms.Label
$labelEndRow.Text                = "to"
$labelEndRow.Width               = 20
$labelEndRow.Height              = 20
$labelEndRow.Anchor              = 'top,left'
$labelEndRow.Location            = New-Object System.Drawing.Point(210,73)

$numEndRow                       = New-Object System.Windows.Forms.NumericUpDown
$numEndRow.Value                 = $null
$numEndRow.Width                 = 100
$numEndRow.Height                = 20
$numEndRow.Maximum               = $null
$numEndRow.Anchor                = 'top,left'
$numEndRow.Location              = New-Object System.Drawing.Point(230,70)
$numEndRow.Enabled               = $false

$labelInclusive                  = New-Object System.Windows.Forms.Label
$labelInclusive.Text             = "(inclusive)"
$labelInclusive.Width            = 60
$labelInclusive.Height           = 20
$labelInclusive.Anchor           = 'top,left'
$labelInclusive.Location         = New-Object System.Drawing.Point(340,73)

$listOutput                      = New-Object System.Windows.Forms.ListBox
$listOutput.Width                = 580
$listOutput.Height               = 240
$listOutput.Anchor               = 'top,left,right,bottom'
$listOutput.Location             = New-Object System.Drawing.Point(10,110)

$btnStart                        = New-Object system.Windows.Forms.Button
$btnStart.Text                   = "Start import"
$btnStart.Width                  = 120
$btnStart.Height                 = 30
$btnStart.Anchor                 = 'right,bottom'
$btnStart.Location               = New-Object System.Drawing.Point(470,360)
$btnStart.Add_Click( { 
    If( $Script:Csv -ne $null ) { 
        If( $Script:Environments[ $textEnvironment.SelectedIndex ].Credential -ne $null ) { 
            $listOutput.Items.Clear()
            Run-Import 
        } Else {
        $listOutput.Items.Add( "No credential provided for $( $Script:Environments[ $textEnvironment.SelectedIndex ].Name )." )
        }
    } Else {
        $listOutput.Items.Add( "No file selected." )
    }
} )

$btnClose                       = New-Object system.Windows.Forms.Button
$btnClose.Text                  = "Close"
$btnClose.Width                 = 60
$btnClose.Height                = 30
$btnClose.Anchor                = 'bottom,left'
$btnClose.Location              = New-Object System.Drawing.Point(10,360)
$btnClose.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton              = $btnClose

$Form.Controls.AddRange( @( $btnFile, $labelEnvironment, $textEnvironment, $btnLogin, $labelStartRow, $numStartRow, $labelEndRow, $numEndRow, $labelInclusive, $listOutput, $btnStart, $btnClose ) )
$Form.ShowDialog()

Function Run-Import() {
    $ServiceUrl = $Script:Environments[ $textEnvironment.SelectedIndex ].ServiceUrl
    $Database = $Script:Environments[ $textEnvironment.SelectedIndex ].Database
    $Credential = $Script:Environments[ $textEnvironment.SelectedIndex ].Credential

    If( $numStartRow.Value -gt $Script:Csv.Length ) { $numStartRow.Value = $Script:Csv.Length }
    If( $numEndRow.Value -gt $Script:Csv.Length ) {}

    For( $RowCount = $numStartRow.Value - 1; $RowCount -le $numEndRow.Value - 1; $RowCount++ ) { # Converting 1-indexed parameter to 0-index
        $Row = $Script:Csv[$RowCount]

        $listOutput.Items.Add( "Processing row $( $RowCount + 1 )" )

        # Determine the System record ID column
        $RowId = $Row.( ( $Row | Get-Member | select Name | ?{ $_.Name -like "*System record ID" } ).Name )
        If( $RowId -imatch '^[0-9a-z]{8}(?:-[0-9a-z]{4}){3}-[0-9a-z]{12}$' ) {
            # a) Get the DataForm

            $LoadRequest = [xml]"<?xml version=`"1.0`" encoding=`"UTF-8`"?>
                                <soap:Envelope xmlns:soap=`"http://www.w3.org/2003/05/soap-envelope`" >
                                    <soap:Header/>
                                    <soap:Body>
                                        <DataFormLoadRequest xmlns=`"Blackbaud.AppFx.WebService.API.1`">
                                            <ClientAppInfo
                                                REDatabaseToUse=`"$Database`" 
                                                ClientAppName=`"TaxDeclarationUpdateTool`" 
                                                TimeOutSeconds=`"60`"
                                                RunAsUserID=`"00000000-0000-0000-0000-000000000000`">
                                            </ClientAppInfo>
                                            <FormID>c5c6eeee-412e-47e7-928b-7cf07366812b</FormID>
                                            <RecordID>$RowId</RecordID>
                                        </DataFormLoadRequest>
                                    </soap:Body>
                                </soap:Envelope>"

            Try {
                $LoadResponse = Invoke-WebRequest -Uri $ServiceUrl -Credential $Credential -Method Post -Body $LoadRequest -ContentType 'application/soap+xml;charset=UTF-8'
            } Catch {
                $listOutput.Items.Add( "Error loading data from Blackbaud CRM ($( $_.ErrorDetails)). Did you use the correct credential for the selected environment?" )
                Return
            }

            $ResponseContent = [xml]$LoadResponse.Content
            $Values = $ResponseContent.Envelope.Body.DataFormLoadReply.DataFormItem.Values

            # b) Update DataFormItem

            $Headers = $Row | Get-Member -MemberType NoteProperty | select Name
            $Headers | %{
                $ColumnName = $_.Name <#
                $ColumnName = $Headers[0].Name <#
                #>
                $MappedColumn = ( $Values.fv | ?{ $_.ID -eq $ColumnName } )
                If( $MappedColumn -ne $null ) {
                    $OriginalValue = $MappedColumn.Value

                    $NewValue = $Row.( $ColumnName )

                    $ColumnType = $( Switch( $MappedColumn.Value.type ) {
                        'q1:guid' { 'ms:guid' }
                        'xsd:dateTime' { 'xsd:dateTime' }
                        'xsd:unsignedByte' { 'xsd:unsignedbyte' }
                        'xsd:boolean' { 'xsd:boolean' }
                        'xsd:string' { 'xsd:string' }
                        default { $null }
                    } )

                    If( $ColumnType -eq $null -or $ColumnType -eq 'ms:guid' ) {
                        If( $NewValue -imatch '^[0-9a-z]{8}(?:-[0-9a-z]{4}){3}-[0-9a-z]{12}$' ) {
                            If( $ColumnType -eq $null ) { $ColumnType = 'ms:guid' }
                        } Else {
                            If( $ColumnType -eq 'ms:guid' ) { $listOutput.Items.Add( "Column type is GUID, but value was not a valid GUID." ) }
                        }
                    }

                    If( $ColumnType -eq $null -or $ColumnType -eq 'xsd:dateTime' ) {
                        $Match = ( [regex] '(?i)(0[1-9]|[12][0-9]|3[01])[/-](0[1-9]|1[0-2])[/-](\d{4})' ).Match( $NewValue )
                        If( $Match.Success ) {
                            $NewValue = "$( $Match.Groups[3].Value )-$( $Match.Groups[2].Value )-$( $Match.Groups[1].Value )T00:00:00"
                            If( $ColumnType -eq $null ) { $ColumnType = 'xsd:dateTime' }
                        } Else {
                            If( $ColumnType -eq 'xsd:dateTime' ) { $listOutput.Items.Add( "Column type is dateTime, but value was not a valid date." ) }
                        }
                    }

                    If( $ColumnType -eq $null -or $ColumnType -eq 'xsd:boolean' ) {
                        If( $NewValue -imatch '^(?:yes|true|y|1)$' ) {
                            $NewValue = 'true'
                            If( $ColumnType -eq $null ) { $ColumnType = 'xsd:boolean' }
                        } ElseIf( $NewValue -imatch '^(?:no|false|n|0)$' ) {
                            $NewValue = 'false'
                            If( $ColumnType -eq $null ) { $ColumnType = 'xsd:boolean' }
                        } Else {
                            If( $ColumnType -eq 'xsd:boolean' ) { $listOutput.Items.Add( "Column type is boolean, but value was not a valid boolean." ) }
                        }
                    }

                    If( $ColumnType -eq $null -or $ColumnType -eq 'xsd:unsignedbyte' ) {
                        $Match = ( [regex] '(\d+)' ).Match( $NewValue )
                        If( $Match.Success ) {
                            If( $ColumnType -eq $null ) { $ColumnType = 'xsd:unsignedbyte' }
                        } Else {
                            If( $ColumnType -eq 'xsd:unsignedbyte' ) { $listOutput.Items.Add( "Column type is unsigned byte (a whole number), but value was not a valid unsigned byte." ) }
                        }
                    }

                    If( $ColumnType -eq $null -or $ColumnType -eq 'xsd:string' ) {
                        If( $ColumnType -eq $null ) { $ColumnType = 'xsd:string' }
                    }

                    If( $ColumnType -ne $null ) {
                        If( $NewValue -ne $MappedColumn.Value.'#text' ) {
                            $listOutput.Items.Add( "Updating $ColumnType column `"$ColumnName`" from `"$($MappedColumn.Value.'#text')`" to `"$($NewValue)`"" )

                            $ns_xsi = "http://www.w3.org/2001/XMLSchema-instance"                             $ns_xsd = "http://www.w3.org/2001/XMLSchema"                            $ns_bb = "bb_appfx_dataforms"
                            $ns_ms = "http://microsoft.com/wsdl/types/"
                            $el = $ResponseContent.CreateElement( "Value", $ns_bb )
                            If( $ColumnType -like 'ms:*' ) { $el.SetAttribute( "xmlns:ms", $ns_ms ) }
                            $el.Attributes.Append( $ResponseContent.CreateAttribute( "xsi", "type", $ns_xsi ) ).Value = $ColumnType
                            $el.InnerText = $NewValue

                            If( $MappedColumn.Value -eq $null ) {
                                $MappedColumn.AppendChild( $el ) | Out-Null
                            } Else {
                                $MappedColumn.ReplaceChild( $el, $MappedColumn.Value )
                            }
                        }
                    } Else {
                        $listOutput.Items.Add( "Could not validate new value for column ""$ColumnName""." )
                    }
                }
            }

    #    c) Save DataFormItem

            $SaveRequest = [xml]"<?xml version=`"1.0`" encoding=`"UTF-8`"?>
                                <soap:Envelope xmlns:soap=`"http://www.w3.org/2003/05/soap-envelope`">
                                    <soap:Header/>
                                    <soap:Body>
                                        <DataFormSaveRequest xmlns=`"Blackbaud.AppFx.WebService.API.1`">
                                            <ClientAppInfo
                                                REDatabaseToUse=`"79599S`" 
                                                ClientAppName=`"TaxDeclarationUpdateTool`" 
                                                TimeOutSeconds=`"60`"
                                                RunAsUserID=`"00000000-0000-0000-0000-000000000000`" 
                                                ClientUICulture=`"en-US`" 
                                                ClientCulture=`"en-US`" 
                                                TimeZone=`"Eastern Standard Time`">
                                            </ClientAppInfo>
                                            <FormID>c5c6eeee-412e-47e7-928b-7cf07366812b</FormID>
                                            <ID>$RowId</ID>
                                            <DataFormItem>
                                                <Values xmlns=`"bb_appfx_dataforms`" />
                                            </DataFormItem>
                                        </DataFormSaveRequest>
                                    </soap:Body>
                                </soap:Envelope>"
            
            $NewDataFormValue = $SaveRequest.ImportNode( $Values, $true )
            $SaveRequest.Envelope.Body.DataFormSaveRequest.DataFormItem.ReplaceChild( $NewDataFormValue, $SaveRequest.Envelope.Body.DataFormSaveRequest.DataFormItem.Values )
            Try {
                $SaveResponse = Invoke-WebRequest -Uri $ServiceUrl -Credential $Credential -Method Post -Body $SaveRequest -ContentType 'application/soap+xml;charset=UTF-8'
            } Catch {
                $listOutput.Items.Add( "Error saving data to Blackbaud CRM. Did you use the correct credential for the selected environment?" )
                Return
            }
        }
    }

    $listOutput.Items.Add( "Finished." )
}
