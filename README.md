Tax Declaration Update Tool
===========================

Loads data for tax declarations from a csv file and updates records via the API.

Please be aware that it makes 2 API calls PER LINE of data in the csv file, so
may take a long time to load a large file and also slow down the Blackbaud 
environment. It may be a good idea to break the load into chunks or run the job
during a quiet period.

To use the tool
---

 1. Prepare the csv file:
    * There must be exactly 1 column with a title that ends in "System record ID" 
      that contains the ID (NOT the Lookup ID) of the tax declaration record to be updated.
    * Each column to be mapped must exactly match the column ID that it corresponds to in 
      the DataForm "Tax Declaration Edit Form 2". These IDs can be found in the DataForm 
      metadata using Design Mode. They are, at the time of writing, detailed in the table 
      at the end of this description. 
    * Each column to be mapped must be valid for the data type (in the table above). Currently 
      the tool can update Guid, Date, String, Integer and Boolean types.
    * The tool does not map values to ids for Code Tables, Simple Data Lists or Value Lists. 
      You will have to look these up and resolve them in the csv file before importing.
 2. Run the tool by right-clicking the file TaxDeclarationUpdateTool.ps1 and selecting "Run with PowerShell".
 3. Click the button "Choose file to import..." and select your csv file.
 4. Select the Blackbaud CRM environment you want to load the data into.
 5. Click the "Credential" button and enter your username and password for the selected environment.
 6. If you want to, you can limit which rows are loaded by using the "Import rows from" and "to" numeric fields.
 7. When you are ready, click the "Start import" button. Any output will be logged to the big listbox.

Tax Declaration Edit Form 2 fields (as of 24-Mar-2020)
---

| Field ID                       | Caption                  | Data type  | Descriptor                                                                |
|---|---|---|---|
| CONSTITUENTID                  |                          | Guid       |                                                                           |
| DECLARATIONMADE                | Made	                 | Date       |                                                                           |
| DECLARATIONSTARTS              | Start date               | Date       |                                                                           |
| DECLARATIONENDS                | End date                 | Date       |                                                                           |
| DECLARATIONINDICATORCODE       | Indicator                | TinyInt    | Value List                                                                |
| DECLARATIONSOURCECODEID        | Source                   | Guid       | Code Table (Declaration Source)                                           |
| CHARITYCLAIMREFERENCENUMBERID  | Reference number         | Guid       | Simple Data List (Charity Claim Reference Number By ID Simple Data List)  |
| SCANNEDDOCSEXIST	              | Scanned documents exist  | Boolean    |                                                                           |
| CONFIRMATIONSENT	              | Sent                     | Date       |                                                                           |
| CONFIRMATIONRETURNED	          | Returned                 | Date       |                                                                           |
| PAYSTAXCODE	                  | Pays tax                 | TinyInt    | Value List                                                                |
| TAXSTATUSCODEID	              | Status                   | Guid       | Code Table (Tax Status)                                                   |
| COMMENTS	                      | Comment                  | String     |                                                                           |
  
