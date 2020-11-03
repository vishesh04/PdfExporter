### PDF Exporter

Azure cloud function to export a file to pdf. Following file types are supported - 
- doc, docx, epub, eml, htm, html, md, msg, odp, ods, odt, pps, ppsx, ppt, pptx, rtf, tif, tiff, xls, xlsm, xlsx

Following APIs are available - 

1. GET /pdf-export?source&name

      `source`: Public url of the file to be converted

      `name`: file name with correct file extension

   Response:

   ```
   {
     pdfUrl: temporary url to download the exported pdf file,
     tempFileId: string
   }
   ```
   `tempFileId` in the response above is an internal id taht is used to identify the original file. original file is uploaded to OneDrive for the export to work. Follwing `DELETE` API should be used to delete this file after the exported pdf has been downloaded

2. DELETE /pdf-export/{id}
  
      `id`: `tempFileId` from the response of the API above

## How it works
In the GET API, original file is first uploaded to OneDrive. This is done by streaming the file through this app to OneDrive. After upload, OneDrive API is used to get the pdf url of the file

In the DELETE API, file uploded to OneDrive for conversion is deleted from OneDrive.

### Deploying
1. Create an OneDrive OAuth app and get the following values to be configured as azure fucntion settings - 

- GRAPH_TENANT_ID
- GRAPH_CLIENT_ID
- GRAPH_CLIENT_SECRET

Note: Authentication to OneDrive is done using OAuth `client_credentials` flow

1. Create a folder in OneDrive and get the DriveId and FolderId. Use these as the values for following settings - 

- GRAPH_DRIVE_ID
- GRAPH_WORKING_FOLDER_ID

3. Delploy as azure cloud function

    
