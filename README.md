# Export-OneNote

PowerShell script for extracting OneNote online notebooks to disk

- Downloads notebook data, including:
  - Structure
  - Pages
  - Images
  - Attached files
- Pages are processed into self contained html files
- Creates a simple html preview of the notebook for verifying the export
- Supports incremental updates
- Result is suitable for further processing or import into another system
- Has been used to export notebooks with 1k+ sections and 10k+ pages(!)


## Usage

- Read prerequisites below
- Set required configuration at top of script
- Select stages to run
  - For smaller notebooks, run all stages in a single go
  - For larger notebooks it can be better to run one stage at a time
  - All stages except for 1 can run incrementally, only downloading missing files
  - More details about stages is below
- Run the script
  - It is common for requests to timeout or fail after a while (throttling). Run 
    again to fetch missing data! Typically you run stage 1 one time. Then the rest 
    a few times to get everything.
- Fetching updates
  - Run everything again to update changes since last run. Default is  to "sync"
    the content (i.e. removed notebook content is also removed from export)
  - Remember that stage 1 must run again to fetch new structure and pages


## Prerequisites

- Install PnP PowerShell: https://pnp.github.io/powershell
- On the PnP app registration:
  - Add the following permissions:
    - Microsoft Graph --> `Notes.Read`
    - SharePoint --> `Notes.Read.All`
  - Don't forget to "Grant admin consent"
  - Copy the application (client) id
- Get URL for the notebook to export
  - This is an MS Graph path. Possible formats:
    - `me/onenote/notebooks/{notebookId}`
    - `users/{userId or userPrincipalName}/onenote/notebooks/{notebookId}`
    - `groups/{groupId}/onenote/notebooks/{notebookId}`
    - `sites/{siteId}/onenote/notebooks/{notebookId}`
  - Here is one way to get a valid URL:
    - Open the SharePoint site containing the notebook
    - In the browser console, enter `_spPageContextInfo.groupId` to get the Group ID
    - Go to Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer
    - Log in
    - Query the following URL (replace `GROUPID` with the one you got above):
      `https://graph.microsoft.com/v1.0/groups/GROUPID/onenote/notebooks`
    - Look through the result to find the correct notebook (look at the displayName).
    - Copy the "self" value. 


## Description of Stages and output structure

1: Fetches the notebook data, all sectionsGroups and all sections
  - Output:
        `/notebook.JSON`
        `/sectionGroups/{group-id}.JSON`
        `/sections/unprocessed/{section-id}.JSON`
  - Moves removed section groups and sections to `/removed`
2: Get page JSON data for all sections in `/sections/unprocessed`
  - Output:
        `/pages/{page-id}/{page-id}.JSON`
  - When a section is done, it is moved: `/sections/unprocessed/{section-id}.JSON` --> `/sections/{section-id}.JSON`
  - Moves removed pages to /removed
3: Get page HTML content
  - Process each page folder, if HTML file does not exist it is downloaded
  - Output: `/pages/{page-id}/{page-id}.HTML`
4: Parse html files.
    - Only if `/pages/{page-id}/{page-id}.PARSED.html` does not exist
    - First it downloads all linked images and files
    - When files have been downloaded, html is re-written:
      - Change links to locally downloaded files instead
      - Change attached files from `<object>` to file icon
      - Add title and date
    - Output:
        `/pages/{page-id}/{resource-name}.ext`
        `/pages/{page-id}/{page-id}.PARSED.html`
5: Finishing tasks
  - Compiles groups, sections and pages into single JSON files with most useful data. 
    (This is duplicate data, but makes it easier if you want to further parse the data.)
  - Output:
      `/sectionGroups.JSON`
      `/sections.JSON`
      `/pages.JSON`
  - Lists pages with "Ink" (since these are not fetched by the script).
  - Generates a simple preview of the notebook content: `/preview.html`


## Known Issues

- You may be throttled by the server after a while. If timeouts starts occurring try CTRL+C and
  wait 30 minutes before running again. The script is designed to automatically resume.
- Large notebooks can take a while to process. Be patient if the script seems to be stuck.
- Very large sections may fail (timeout) when fetching pages. Try splitting the section into multiple?
- If pages are not found, it could be that they are not (yet?) synced to the server.
- Moving a page to another section in OneNote counts as a removed page that will be downloaded again
- Tracking removed pages may not work perfectly if stage 2 is interrupted
- Some OneNote features are not included or visible, including:
  - Drawings ("Ink"/"InkML") (can be extracted but not implemented in this script)
  - Loop components (not exported by OneNote)
  - Page background colors and grids (not exported by OneNote)
  - Images are downloaded in full size, not their preview
  - Only single tags (i.e. not multiple) are visible in "parsed" html
  - Some images are returned as an "octet-stream" mime type. All such turned out to be EMF images 
    for me. Browsers can't display EMF. So the script will fetch the preview images for these instead. 
    All octet-stream images are assumed to be EMF and downlaoded too. I can't promise that this is 
    correct in all cases.
  - There may be other kind of content I have not seen yet


## Useful Links:

- GET: https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/concepts/onenote-get-content.md
- HTML: https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/concepts/onenote-input-output-html.md
