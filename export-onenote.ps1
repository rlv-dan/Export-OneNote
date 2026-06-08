<#

# Export-OneNote

https://github.com/rlv-dan/Export-OneNote


## PowerShell script for extracting OneNote Online data

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
- Set required configuration below
- Select stages to run
  - For smaller notebooks, run all stages in a single go
  - For larger notebooks it can be better to run one stage at a time
  - All stages except for 1 can run incrementally, only downloading missing files
    For full description see below section
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

#>


# --- Required Configuration ---------------------------------------------------------------------------------

# Client ID for the app registration
$appId = "11111111-1111-1111-1111-111111111111"

# Tenant URL
$site = "https://mytenant.sharepoint.com"

# URL to notebook
$notebookUrl = "groups/11111111111111111111111111111111/onenote/notebooks/1-11111111-1111-1111-1111-111111111111"
  
# Where to save the exported data. A folder will be created here
$exportpath = "c:\temp\my_notebook_exports"

# Select stages to run. To skip a stage comment out the line
$stages = 
    1,
    2,
    3,
    4,
    5


# --- Additional options -------------------------------------------------------------------------------------

# When updating: Should removed items in OneNote also be removed from disk?
$detectRemovedItems = $true
# Move removed items to /removed folder instead of deleting them from disk
$keepRemovedItems = $true

# Apply a filter when fetching pages
#   Typically used to fetch only pages updated since a certain date (faster)
#   Cannot be combined with $detectRemovedItems
#   Example: "&filter=lastModifiedDateTime+ge+2025-06-01"
$pageFilter = ""

# Pausing between requests might reduce risk of throttling.
$requestPauseInMilliSeconds = 1000

# Default timeout is 100 seconds. This may happen for large notebooks or when being throttled
# Timeout can be increased here, but it does not seem to help loading large items. I think this
# is because the timeout is at the Microsoft gateway side.
#$env:SharePointPnPHttpTimeout = 30


# ------------------------------------------------------------------------------------------------------------

if($pageFilter -ne "") {
    $detectRemovedItems = $false
}

function Get-Percent($current, $max) {
    if($max -eq 0) { return 0 }
    return [Math]::Clamp(($current / $max) * 100, 0, 100)
}

function EnsurePath($path) {
    if (!(Test-Path -path $path)) {
        New-Item -ItemType Directory -Force -Path $path
    }
}

EnsurePath("$exportpath")
EnsurePath("$exportpath\sections")
EnsurePath("$exportpath\sections\unprocessed")
EnsurePath("$exportpath\sectionGroups")
EnsurePath("$exportpath\pages")

function Get-SectionGroupsRecursive($root) {
    $result = @()
    Write-Host "$($root.displayName)..."
    $groups = Invoke-PnPGraphMethod -Url $root.sectionGroupsUrl -All
    Start-Sleep -Milliseconds $requestPauseInMilliSeconds
    $result += $groups.value
    foreach($group in $groups.value) {
        $group | ConvertTo-Json -Depth 100 | Set-Content -Path "$exportpath\sectionGroups\$($group.id).json" -Encoding utf8 -Force
        if($group.sectionGroupsUrl) {
            $result += Get-SectionGroupsRecursive($group)
        }
    }
    return $result
}

function Remove-DeletedSectionGroups($currentGroups) {
    $groupsOnDisk = Get-ChildItem -Path "$exportpath\sectionGroups\*.json" | Select-Object -ExpandProperty "Name"
    $currentGroupsIds = $currentGroups | Select-Object -ExpandProperty "id"
    foreach($filename in $groupsOnDisk) {
        $id = $filename.Replace(".json","")
        if($currentGroupsIds -notcontains $id) {
            write-host "REMOVED GROUP: " $id
            if($keepRemovedItems -eq $true) {
                EnsurePath "$exportpath\removed\sectionGroups" | Out-Null
                Move-Item -Path "$exportpath\sectionGroups\$filename" -Destination "$exportpath\removed\sectionGroups\$filename" -Force
            }
            else {
                Remove-Item -Path "$exportpath\sectionGroups\$filename" -Force
            }
        }
    }
}


function Get-Sections($groups) {

    $sectionsOnDiskBefore = Get-ChildItem -Path "$exportpath\sections\*.json" | Select-Object -ExpandProperty "Name"
    $sectionsFetched = @()

    foreach($group in $groups) {
        Write-Host "$($group.displayName)..."
        $groupSections = Invoke-PnPGraphMethod -Url $group.sectionsUrl -All
        if($null -ne $groupSections -and $null -ne $groupSections.value) {
            foreach($section in $groupSections.value) {
                $section | ConvertTo-Json -Depth 100 | Set-Content -Path "$exportpath\sections\unprocessed\$($section.id).json" -Encoding utf8 -Force
                $sectionsFetched += $section.id
            }
        }
        Start-Sleep -Milliseconds $requestPauseInMilliSeconds
    }

    if($detectRemovedItems -eq $true) {
        write-host "Purge removed sections..." -ForegroundColor Cyan
        foreach($filename in $sectionsOnDiskBefore) {
            $id = $filename.Replace(".json","")
            if($sectionsFetched -notcontains $id) {
                write-host "REMOVED SECTION: \sections\" $filename
                if($keepRemovedItems -eq $true) {
                    EnsurePath "$exportpath\removed\sections" | Out-Null
                    Move-Item -Path "$exportpath\sections\$filename" -Destination "$exportpath\removed\sections\$filename" -Force
                }
                else {
                    Remove-Item -Path "$exportpath\sections\$filename" -Force
                }
            }
        }
    }

}

function Get-PageJSON() {

    $pagesOnDiskBefore = Get-ChildItem -Path "$exportpath\pages\*" -Directory | Select-Object -ExpandProperty "Name"
    $pagesFetched = @()
    $sectionsProcessed = @()

    $sections = Get-ChildItem -Path "$exportpath\sections\unprocessed\*.json" | ForEach-Object { (Get-Content -Raw $_.FullName | ConvertFrom-Json) }
    $count = 0

    foreach($section in $sections) {
        $count++
        Write-Progress -Activity "Get JSON for pages in each section" -Status $section.displayName -PercentComplete (Get-Percent $count $sections.Count)
        $url = $section.pagesUrl
        if($pageFilter -ne "") {
            $url = $url.Replace("/v1.0/","/beta/") # using filter requires beta endpoint
        }
        $sectionPages = Invoke-PnPGraphMethod -Url "$url/?pagelevel=true$pageFilter" -All
        if($null -ne $sectionPages -and $null -ne $sectionPages.value) {
            foreach($page in $sectionPages.value) {
                $id = $page.id
                $pagesFetched += $id
                if((Test-Path "$exportpath\pages\$id\") -eq $true) {
                    $existingJson = Get-Content -Raw "$exportpath\pages\$id\$id.json" | ConvertFrom-Json
                    $newJson = $page | ConvertTo-Json | ConvertFrom-Json
                    if($existingJson.lastModifiedDateTime -ne $newJson.lastModifiedDateTime) {
                        Write-Host "UPDATED PAGE:" $page.title
                        Remove-Item "$exportpath\pages\$id\*" -Recurse -Force | Out-Null  # clear this page so that everything is downloaded again
                        $page | ConvertTo-Json -Depth 100 | Set-Content -Path "$exportpath\pages\$id\$id.json" -Encoding utf8 -Force
                    } else {
                        # Write-Host "ALREADY UP TO DATE:" $page.title
                    }
                } else {
                    # new page
                    Write-Host "NEW PAGE:" $page.title
                    EnsurePath("$exportpath\pages\$id\") | Out-Null
                    $page | ConvertTo-Json -Depth 100 | Set-Content -Path "$exportpath\pages\$id\$id.json" -Encoding utf8 -Force
                }
            }
            Move-Item -Path "$exportpath\sections\unprocessed\$($section.id).json" -Destination "$exportpath\sections\$($section.id).json" -Force
            $sectionsProcessed += $section.id
        }
    }

    if($detectRemovedItems -eq $true) {
        write-host "Purge removed pages..." -ForegroundColor Cyan
        foreach($filename in $pagesOnDiskBefore) {
            if($pagesFetched -notcontains $filename) {
                $id = $filename.Replace(".json", "")
                $existingJson = Get-Content -Raw "$exportpath\pages\$id\$id.json" | ConvertFrom-Json
                if($sectionsProcessed -contains $existingJson.parentSection.id) {
                    write-host "REMOVED PAGE: \pages\$filename"
                    if($keepRemovedItems -eq $true) {
                        EnsurePath "$exportpath\removed\pages" | Out-Null
                        Move-Item -Path "$exportpath\pages\$filename" -Destination "$exportpath\removed\pages\$filename" -Force
                    }
                    else {
                        Remove-Item -Path "$exportpath\pages\$filename" -Force
                    }
                }
            }
        }
    }

    Write-Progress -Completed
}

function Get-PageHTML() {
    $pageIds = Get-ChildItem -Path "$exportpath\pages" -Directory | select-object -ExpandProperty  "Name"
    $count = 0

    foreach($id in $pageIds) {
        $count++
        if((Test-Path "$exportpath\pages\$id\$id.html") -eq $false) {
            $page = Get-Content "$exportpath\pages\$id\$id.json" -raw | ConvertFrom-Json
            Write-Progress -Activity "Get Page HTML" -Status ($page.title -eq "" ? "Untitled Page" : $page.title) -PercentComplete (Get-Percent $count $pageIds.Count)
            $pageContent = Invoke-PnPGraphMethod -Url $page.contentUrl -Raw
            if($pageContent -ne "" -and $null -ne $pageContent) {
                $pageContent | Set-Content -Path "$exportpath\pages\$id\$id.html" -Encoding utf8 -Force
            } else {
                Write-Host "Skip empty or null page: $($page.title) [$($id)]" -ForegroundColor Red
            }
            Start-Sleep -Milliseconds $requestPauseInMilliSeconds
        }
    }
    Write-Progress -Completed
}

function Get-PageLinkedResources {
    $pageIds = Get-ChildItem -Path "$exportpath\pages" -Directory | select-object -ExpandProperty  "Name"
    $titleHTMLRaw = '
        <div style="position:absolute; left: 48px; top: 0px; z-index: 1; white-space: nowrap;">
            <h1 style="border-bottom: solid 1px #c8cacc; margin-bottom: 6px;">{{PAGETITLE}}</h1>
            <div style="color: gray;">{{PAGEDATE}}</div>
        </div>'

    $fileIconRaw = '
            <a style="text-decoration: none;" href="{{FILENAME}}">
                    <div style="width: 72px; height: 92px; overflow: hidden; text-overflow: ellipsis; text-align: center;">
                        <div style=''width: 48px; height: 48px; margin-left: auto; margin-right: auto; background-image: url("data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyMDQ4IDIwNDgiIGNsYXNzPSJzdmdfYzY5NmI0YTYiPjxwYXRoIGQ9Ik0xNzkyIDU0OXYxNDk5SDEyOFYwaDExMTVsNTQ5IDU0OXptLTUxMi0zN2gyOTNsLTI5My0yOTN2Mjkzem0zODQgMTQwOFY2NDBoLTUxMlYxMjhIMjU2djE3OTJoMTQwOHoiPjwvcGF0aD48L3N2Zz4=");''>
                        <span style="display: inline-block; color: white; background-color: {{EXTCOLOR}}; width: 100%; font-size: 12px; font-weight: bold; margin-top: 26px; margin-left: -2px;">{{FILEEXT}}</span>
                    </div>
                    <div style="overflow: hidden; font-size: 12px; text-overflow: ellipsis; margin-top: 2px;">{{FILENAME}}</div>
                </div>
            </a>'

    $tagStyling = "
        <style>
            *[data-tag] {
                position: relative;
            }
            *[data-tag]:before {
                position: absolute;
                left: -1.25rem;
                top: 0.1rem;
                width: 16px;
                height: 16px;
            }
            *[data-tag='to-do']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAySURBVDhPY5TrPvufgQIANuBOvi6USxpQmXiZgQnKJhuMGjBqAAiMGgDNjVA2GYCBAQAn7Ap9ZukRFwAAAABJRU5ErkJggg==')}
            *[data-tag='to-do:completed']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACbSURBVDhPY5TrPvufgQIANuBOvi6USxpQmXiZgQnKJhvQ14A/R44wfCssgPIggGgDQJp/TprAwBYcAhWBAKIMgGlmzytgYLGxgYpCAIoBIOeBFCMDfJpBAMUAkPNAimGGENIMAixQGgxgikCa/t27y/B7y2a8mkEAIwxAikGaiNEMAlgDEaSJe9UagppBgKhYwAcG3gAKszMDAwBGoUBdvU14NgAAAABJRU5ErkJggg==')}
            *[data-tag='important']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAGgSURBVDhPY8AHLC0tDS0sLAKhXKyACUpjBXx8PMXy8jLFUC5WwAilMYCZmZkMNzfn3Z8/f7L9+fNP9tSpU0+gUigApwuYmRmDgM7/7+jo+FdWVjoFKowBcBogIMBf4eLixu7k5ML8+/fvHKgwBmC0tDTPY2RknAjlw4GwsPDftWvXM4PYwcGBf9++fQtmo4L/y8EUyJCQkKCvT58+/f/r1y+8GKQmPDz0l6ur00awZhggxhCcmmEAnyEENcMA0JCaxMSEb+gGFBTk/8GmGWssSEiIY4grKSkxc3Nzc0G5cIChkJeXt8TfP4AdxD58+DAYg4CzswvD589fHMAcJIASNZaWZk5AA+Kjo2NZpk6d8mPZsiVfDx488OXs2bP/TU3N2B49evSHiYnx05MnT09CtaAaICsrUyYuLm62ZMniv0+ePF794cOnoCNHjtaysDAJHDiwX4+Li4uDlZVV5/r16xOgWlDzgpOTw2tmZubvX758Tjh+/NQ+qDAYgHImUH/v9+/fHf//ZzQ6fvz4eagUBICcD4oBKBcnAGVvhDoGBgDGmfepDiBZxwAAAABJRU5ErkJggg==')}
            *[data-tag='question']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAEISURBVDhPxZG/agJBEMZn7u4d0qfIQwhpNF5OBFOGq+0FWx8gD2AlBOvTLieIItiJdeqIjU8g2HqOM+t45DZ6GBT8NfPn7vt2ZhfBIioNn8BNWvwh4PLBNAnmhNAOJ7W+qX+RMRAxutuR5AT46SCsJd/t6AURAu7VbRNP4wE+WQIlXiWcVn9M70An8uMxEjQ4zxg4Gg0I9Myn9C2xwUH84h8KWqZkDNjikcdeaXFbIn/w0fPjpZYp1gSnYfE7r9eQi9VWyp9ntFFxlwBm4eRNnjZDrsHxWVm8OCUW8ldwt20i2JwTC7kGCPiKSE0t/4eM3/MHPH0+ZydwvKSo6Z2IyvH3VStcBsAelClSbvcbDiUAAAAASUVORK5CYII=')}
            *[data-tag='remember-for-later']:before { content: ''; background-color: #FFFF00; width: 1rem; height: 1rem; margin-top: 0.25rem;}
            *[data-tag='definition']:before { content: ''; background-color: #00FF00; width: 1rem; height: 1rem; margin-top: 0.25rem;}
            *[data-tag='highlight']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAEQSURBVDhPY2QgEVhamuexs7N3g9g/f/4sBQsSC0CanZwc/pw4cfw/CDs42P1kgsoRBCDNnJycfW1t7cxGRsZQUSIBss2/fv0C2w7ig8ShSnCDIar5ra11w01H+7/ntm4hqJkZSsMBSDMDG1sNr6YWs8CJ4wy3eHgYylua/37//r3o+PGTk6DKsAOQ5rfOjn++7dsHtvlDdtZ/kEuIc7aNVRGyZhAN4oNdRAx46+L08evGjeRpfmNr3fs+Ie43SDPIEFI0M36wslL+y8Z8hcXegePfw4d//j158o3hx49G4SPH+qBq8AJGoN9vMTAyqjIwM/1g+PuvU/jwUeKcDQWQzPT//23mX391SNXMwMDAAADCgdsCOcg9CQAAAABJRU5ErkJggg==')}
            *[data-tag='contact']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADhSURBVDhPYxj6gBFKM7i7uy759etXKIjNxsa2+v37jxWnTp16ApbEA8AGgDSbmJhEZGXlMIP406ZN+XvmzJkVO3fujgHxraws/oNodHDs2AmIAxwd7X8+ffr0P9AFYAxig8TAkgQAE5QmG7CACJCfgc5G8QJIDMQGAXxeAPvB0tLSkJWV+cifP3+4QHwWFpZvv3//tTl+/Ph5EB8vsLQ0c3JwsPs5Z85ssN9BGMQGiYHkoMpwAhZBQeHWoKAgtri4eKgQAwOUzbZu3bpWIG2JNxZAksgxgBwTuDSOAmTAwAAAJp59rOGvW3QAAAAASUVORK5CYII=')}
            *[data-tag='address']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAC6SURBVDhPY8AHLC3NnEAYysUKmKE0BgBpFBYW2cbFxRUtLCx47MmTp/ehUoQBSLOPj9ePEyeO/wdhEJuQS+AAWfOvX7/AmGhDsGkm2hB8mmEYpyHEaIZhDENADEdH+5/EaIZhkFqQHpBeJgYGRo+fP3+yGRkZgw0kBoDUgvSA9DKCBKysLP4fOHAILAkDDg52UBYEYJM/duwEI9AFuAFIAQhDuVgBXgOIAcPAAHgsgHloABaAuOVPMAIA6ejlzDtSbQQAAAAASUVORK5CYII=')}
            *[data-tag='phone-number']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAH6SURBVDhPY2QAAgsLi0A+Pp75X7584RcTE93z7NnLxFOnTj0ByRECTCDN3Nycq8rLK/lXr17LoK2t4ygqKrIbKk8YBAT4n1u+fNn/X79+wXFoaPB3KyvzBKgSvIDp1auXhnZ29lAuBHBz8/z994/xI5SLFzCBCFFRUTAHBM6dO8tw//49ln///p2GCuEFTGxsbL9u374N5TIwzJw58zvQG2VEB6KAAP+hS5cuQrkMDDdv3uBkYPh/BcolCJiFhASZP3785O7r68sKEuDl5WW4cOFCpKSk+LEnT57eB6vCA5gfP356QVJSLFtQUJBXRUWVQUtLi4GLi4uZWEOYQYSoqPD5c+fORbu7ezBxc3PjNMTMzExGXl4mSFZWxgBkMUgMnBJBwM/Pp4eTkzOvr28CKyxW1q5dAwzUGb9+/vzhCVSqw87O3q2jo/vv69ev/1+9evXw9es3rmCFMODq6rQxPDz019OnT+GJCpTIrKws/icmJnw7ceI4XLy6uvKPu7vrEqhWBMBmCDYMkgcZDE5IyGD37n3+wEw1Ly4u5s+OHduhopjg7NkzDBwcHPfBgYgObt++s0VCQuzolSuX3Q4cOMDy69dPViBgEBYWZnj9+jU8bH78+B4OD0RcAJSpxMTEor9//2H++fNnXpCYgIDgiffv31YfP35qHwBofRB+SGTI0AAAAABJRU5ErkJggg==')}
            *[data-tag='web-site-to-visit']:before, *[data-tag='source-for-article']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyFpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNi1jMTQyIDc5LjE2MDkyNCwgMjAxNy8wNy8xMy0wMTowNjozOSAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENDIChXaW5kb3dzKSIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDpGRDRGRkZFNzU4NzExMUU4OUU3RUVGRjUyRjJEMERDMCIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDpGRDRGRkZFODU4NzExMUU4OUU3RUVGRjUyRjJEMERDMCI+IDx4bXBNTTpEZXJpdmVkRnJvbSBzdFJlZjppbnN0YW5jZUlEPSJ4bXAuaWlkOkZENEZGRkU1NTg3MTExRTg5RTdFRUZGNTJGMkQwREMwIiBzdFJlZjpkb2N1bWVudElEPSJ4bXAuZGlkOkZENEZGRkU2NTg3MTExRTg5RTdFRUZGNTJGMkQwREMwIi8+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+cXK+mQAAATFJREFUeNpi/P//PwMlgImBQsACpW2A2BSI+bGo+QjEp4H4CC4DTKdOnVyyceMmh69fv2AYwM3N89Hf3+9AdnbuT0ZGxtMYBgA1V+/cucMyKip6R0JCwg10BQsWLNBYt261I5BZDQyvHqAhqC5xdXX5MH/+/BVASVNkcSDfBogLgbgBJA9SB+NDaRuwQisri/8gQTTNplOmTNoA0gSTR6ZB4iB5kDoWkB9BzgRxQH4EmQwKE2RvgeRB6kABevTo8UZkb7GAAgjKAdkMCoOPoABF1gySB6kDxQYoDGDeXbZsqQdJzkXzJlieBersHmA0HYSmg48wb4GcC0sHyFEIMgzmLXBCgkbNEVjoI3sLGrUuQHEX9KgFqWPElrpANoDSBxGJq5URT2YiKnkzDnhuBAgwAMVyBMT2pFffAAAAAElFTkSuQmCC')}
            *[data-tag='idea']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAFPSURBVDhPtZKxToNQGIWvFSUdsBNpcDZxcWqaKgwYi32N+gbupomriw9gta6O9g0cJCkEYh+g16mRgREwFtAE/W9/b9pyOzj4JeRyzn/u4RLYICvoeqtdr2vXcRwdpGm6JctyXqvtPIdheOU43hPGxOj64bllnXwOBndFEARFnudsBQ0+zDFaBp4MIdd12MbVC/x5SauNWxibuJJms3Hb7Z7tWdYpOsto2i6pVquVyYTuU/p6jzap4EqiKDZN8xiVGJjPZh8NlAxekGXZtqqqqMTAHHIoGbxAUZRkPH5BJYZSynIoGbxAkqSHH3KUQvr9mxxyKMt0Ota017v4En0F8GGOUQ4/ARBFieH7/rtt2+jMAQ0+zNHiLBV4nveWJMnlcPiYosUADT7M0eKUfmXAMI4KvOWMRq4wKwQKFt9fVPjL/5zgL6xtXTzF+qcT8g10zseptbp/3gAAAABJRU5ErkJggg==')}
            *[data-tag='password']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACHSURBVDhPY2TAAiwsLAKZmP43MTAw6kBE/l/594+x7sSJE+shfARggtJwANHMMB+k4dixE4wgDGKDxEByUGVwgGEAIyNDyv///1FsA7FBYiA5qBBuYGVl8R/KxADY5DBcQCoYeAPg0YjP79gAKHagTAgAGfDr1y+iMLJlwyAQh4EBFKYDBgYAqshasLJoBrcAAAAASUVORK5CYII=')}
            *[data-tag='critical']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAoSURBVDhPY8AH3tpanQFhKBcrYILSZINRA0YNAIGBN2CQAzrkRgYGACC3CNn+vfQIAAAAAElFTkSuQmCC')}
            *[data-tag='project-a']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAqSURBVDhPY6AUMIKIF4LK/8E8EoHE+7uMTFA22WDUgFEDQGDUAIoBAwMA9WkEHNNb0TsAAAAASUVORK5CYII=')}
            *[data-tag='project-b']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAqSURBVDhPY6AUMIKIX3tF/oN5JAI25zeMTFA22WDUgFEDQGDUAIoBAwMAG8AEHF+70B4AAAAASUVORK5CYII=')}
            *[data-tag='movie-to-see']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAABKSURBVDhPYxhwwGhlZfEfyiYPgAz49evXfxANY+PDyOpANNgFBw4cghpHGnBwsGNggrLBHBAmBNDVwQ0gF4x6YdB4AcoekoCBAQDoBUuVSQ7DeAAAAABJRU5ErkJggg==')}
            *[data-tag='book-to-read']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACaSURBVDhPY7SysvjPgAbY2dl/sbGxrd65c3eMu7vrkl+/foX+/PmTDSoNB8eOnWBkABkAVICCnz59+r+6uvKPvb3tCxAN4qOrgVnMBDYKDYiKijJkZeUw//79WxxEg/i4AFYDQACmCZ9mEMBpALFg1IBRA0AAa2YCAVlZ2Z+PHz9mh9FQYRQAzky4gJWVeQIoM4FoqBAtAAMDAJAhVu9437tsAAAAAElFTkSuQmCC')}
            *[data-tag='music-to-listen-to']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADmSURBVDhPYyAIkjf+Z0jZqAjlYQAmKI0f/PtfCmVhAOIMYGT0Z0jeUA3loQDiDGD4lwE0JYsheb0vVAAOiDNgbuBmBob/0xj+MyZDReCASBcAASPjMiB2gvLggHgD5vjfZ/j//waURwIARSMeQLwLcICBN4ARSkMAMMlys7Eu/Przty2Iy83OehjMnuuPqg4JoLiAi4Vll7Qwj2WRvzEDCCtJ8FtApXACFAO+/f6jEuuoxSLIw8EAwmE26qxQKZwA1QWsLHcW77/25/2XHwwgDGKzszBdh0pjBYTD4NfveHAiwgoYGAApREKb9ZYtjwAAAABJRU5ErkJggg==')}
            *[data-tag='remember-for-blog']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACBSURBVDhPY6AUMIIIKyuL/2AeieDYsRNg/WADfv36RRKGWcoENoECMGrAYDAAIyEdOHCI4fXr1wxFRQW/Hz9+zAoVxgrgCQkGQAY9ffr0f3h46C9XV6eNUGG8AMUEkAGysrK/3717u3337n3+UGG8ACMMSNGMAayszJdBmfQCDAwAG9JdrSDkoOcAAAAASUVORK5CYII=')}
            *[data-tag='discuss-with-person-a']:before, *[data-tag='discuss-with-person-b']:before, *[data-tag='discuss-with-manager']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADNSURBVDhPY2QAAtnuM51sTEyFIPavf//6H5ealGtMOPft2+//nCAxgkC599yvtVde/wdhEBsk9uvXr/VA/B8flus++58JbAIFAGrA/7uF2x8xgDCIrdB5Sh3I2AeWIgCYNCee32EqxaO6O06NAYRBbE4O1olsbGyTgfKPIcpwA6bff/87NzpJMSsKsjOAMIj99dc/d6j8WSiNE4BjARcABtR6IBUA4WEClYmXGZhAIYkeujAMVINTMwxQKxbIB6MGUMEARlA6gLLJAAwMAAGHfI1ozywsAAAAAElFTkSuQmCC')}
            *[data-tag='discuss-with-person-a:completed']:before, *[data-tag='discuss-with-person-b:completed']:before, *[data-tag='discuss-with-manager:completed']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAEgSURBVDhPY2QAAtnuM51sTEyFIPavf//6H5ealGtMOPft2+//nCAxgkC599yvtVde/wdhEBsk9uvXr/VA/B8flus++58JbAIFAGrA/7uF2x8xgDCIrdB5Sh3I2AeWIgCYNCee32EqxaO6O06NAYRBbE4O1olsbGyTgfKPIcog4M+RIwzfCgugPAhg+v33v3OjkxSzoiA7AwiD2F9//XOHyp+F0mDNPydNYGALDoGKEAGAAQUOyG/79v1/H+AHptEDkQFEwAQ+ZGdhKMKlGYQxYgHkPJAzQc4FAZiz2fMKGFhsbMBi6IAFSoMBTBFI0797dxl+b9mMVzMIYKQDkGKQJmI0gwDWhATSxL1qDUHNIECtlEg+oNgARnBiIBswMAAA5pzZXTA7x+oAAAAASUVORK5CYII=')}
            *[data-tag='send-in-email']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAGjSURBVDhPYzAzM5NhoAAwy8vLzNDQUI8TFxd99eTJ0/tQcaIBs6ysDO/v33/KhYWFIzQ1NROFhAQYgAadhMoTBIwgwsHB7ufUqdPZ7t69w7Bhw4afT58++fXp06cZDAxMy48fP34erBIHYAYR6upqGj9+/NCKi4tn8vHxZTE1NWP/8uWzxcuXL1L19XXtBASEvjx58uQGWAcaALvA0tLSUFhY8Pi6dRvYwaJQ8Pr1a4bt27cBXbX+19+/f95++PCx4+/f/+tOnTr1BKoE4gKg6S9A/ufj4xVUUVEFS4AANzc3g76+AUN4eASzjIws7+fPn53fv3+XD3KxmJj4fZA+sAtAwNLSvIuRkbEUyiUK/P//Px9ugL+/393k5GQlDw9PqAgqOHz4MMPGjRt+Xbp0kYGdnX31p09fekEBTHEYsIAIPj6eYisra3B4gMDt27cZli1b8hdowz9guBx88+bdtBMnTqyHSqMAitMBo5WVeQKQmi8kJPTz/3+G++/evZ1+/PjJSVB5ggBkwDJgdHF/+fJl4vHjp/ZBxYkHlOZGCgEDAwDxFc9UjBCgLAAAAABJRU5ErkJggg==')}
            *[data-tag='call-back']:before, *[data-tag='schedule-meeting']:before  { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAC3SURBVDhPY2RAAgqdp9T/MjId+s/EVMDy9885Hk6WE59+/heASmMFKAaAgGz3mU7G/wzZ/5iYUm9laRxjYWE5DBKGyKIClYmXGRjlus/+h/JRwP///189LjMR//Xrly2QewgiigrABkDZcECqC5hALgDaAsIgm0AmehCjGQaYoDQILP327Zv8o1ITfWI1gwCyAbIgTSCXEKsZBJANAAGQJlCAEaUZBNANIBmMGjAYDMCZmYgDDAwAF39LPaTOFwkAAAAASUVORK5CYII=')}
            *[data-tag='call-back:completed']:before, *[data-tag='schedule-meeting:completed']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAEXSURBVDhPY2RAAgqdp9T/MjId+s/EVMDy9885Hk6WE59+/heASmMFKAaAgGz3mU7G/wzZ/5iYUm9laRxjYWE5DBKGyKIClYmXGRjlus/+h/JRwP///189LjMR//Xrly2QewgiigrABkDZcECqC5hALgDaAsIgm0AmeuDS/OfIEYZvhQVQHhQgGfDo27dv8iAxEA3iQ8XB+Nu+ff/fB/iBaZgYSC8T2BQIkAXZCJSwxWbzz0kTGNjzChhYbGygohCAbADIebJAxaAAI0ozCKAYwBYcAlYM0gQChDSDAAuUBgOYIpCmf/fuMvzeshmvZhBAcQEIgBSDNBGjGQQwDAABkCbuVWsIagYBrAaQAgbeAJyZiTjAwAAAaeym2QOwz+cAAAAASUVORK5CYII=')}
            *[data-tag='to-do-priority-1']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAABjSURBVDhPY5TtOrOIkYEh6FGZCY9c99n/DCQCRqCmaUA681GpMeOvX79IMkBl4mUGJqDGLCifLMAEpckG1DEA5H8wjwzABAp5UOCRGoAwMEjCgBIwagAVDADlRrISEAQwMAAAR0ogSzB4qxkAAAAASUVORK5CYII=')}
            *[data-tag='to-do-priority-1:completed']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADGSURBVDhPY5TtOrOIkYEh6FGZCY9c99n/DCQCRqCmaUA681GpMeOvX79IMkBl4mUGJqDGLCifLMAEpckGJBnw58gRhm+FBVAeBIANAPkfzMMDQJp/TprAwBYcAhWBACZQyIMCD18AwjSz5xUwsNjYQEUhAMULIOeBFCMDfJpBAMUAkPNAimGGENIMAixQGgxgikCa/t27y/B7y2a8mkEAIxZAikGaiNEMAlijEaSJe9UagppBAKsBpICBNwCUG0nOwgjAwAAAZf9WK8xtq2sAAAAASUVORK5CYII=')}
            *[data-tag='to-do-priority-2']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACnSURBVDhPY5TrOrvrP8N/fQYGhq/M//95Kotx191+8yMKyCcOKHSeUgfRsl1nXsp1ncn+9etXLhD/JwbLdZ/9z/Sg3OwmUPMioBlfH5WZTAUZRgpgku0+u46RgVEC5HyoGEmAiZGBIZCBkcH1HxPzDZAXoOJEA6B+VAD0Wy6QmgTh4QcqEy8zMIECAjlggOJEaYYBJihNNhg1YFgYwAhKSFA2GYCBAQB7qmWZ1PSaeAAAAABJRU5ErkJggg==')}
            *[data-tag='to-do-priority-2:completed']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAEDSURBVDhPY5TrOrvrP8N/fQYGhq/M//95Kotx191+8yMKyCcOKHSeUgfRsl1nXsp1ncn+9etXLhD/JwbLdZ/9z/Sg3OwmUPMioBlfH5WZTAUZRgpgku0+u46RgVEC5HyoGEmAiZGBIZCBkcH1HxPzDZAXoOJYwZ8jRxi+FRZAeTgA0G9Yw+Dbvn3/3wf4gWmYGCgMGEAEskJsGJtmEAYHItRiMAA5D+RMZADi/5w0gYE9r4CBxcYGKooAKAawBYeAFcMMIaQZBFigNBjAFIE0/bt3l+H3ls14NYMAigtAAKQYpIkYzSCAYQAIgDRxr1pDUDMIYDWAFDDwBjCCUxPZgIEBAGee4CeEc5JRAAAAAElFTkSuQmCC')}
            *[data-tag='client-request']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAABXSURBVDhPY3xra/3/PwPDbJHDR9Pkus8CmaQBFAN+/fpFkgEqEy8zMEHZZINRA4CxAKXBgKxYAEXjG1vrWVAxksEg8AIo+YI0kqoZBoZDOiAnCyMAAwMACX8wFr1o52cAAAAASUVORK5CYII=')}
            *[data-tag='client-request:completed']:before { content: ''; background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAC4SURBVDhPY3xra/3/PwPDbJHDR9Pkus8CmaQBFAN+/fpFkgEqEy8zMEHZZAP6GvDnyBGGb4UFUB4EEG0ASPPPSRMY2IJDoCIQQJQBMM3seQUMLDY2UFEs4EN21v9v+/b9B8UGDIP47wP8MMRBGBTtTKBofGNrPQtkAMh5IJtANoIAMTazQGkwgCkCafp37y7D7y2bCTsbGcCchs/ZyBjkBUYQcSdfF2oEaWAIpkRsABwGUDYZgIEBANc+oSKonD2iAAAAAElFTkSuQmCC')}
        </style>
    "

    $count = 0
    foreach($id in $pageIds) {
        $count++
        $page = Get-Content "$exportpath\pages\$id\$id.json" -raw | ConvertFrom-Json
        $title = ($page.title -eq "" ? "Untitled Page" : $page.title)
        Write-Progress -Activity "Download Resources" -Status $title -PercentComplete (Get-Percent $count $pageIds.Count)
        $fileSrc = "$exportpath\pages\$id\$id.html"
        $fileParsed = "$exportpath\pages\$id\$id.PARSED.html"

        $html = New-Object -ComObject "HTMLFile"
        $content = Get-Content $fileSrc -raw
        $html.write([System.Text.Encoding]::Unicode.GetBytes($content))
        $imgTags = $html.all.tags("IMG")
        foreach($img in $imgTags) {
            # Download Image
            $fullResSrc = $null
            $mediatype = $null
            $emfSrc = $null
            try {
                $fullResSrc = $img.attributes["data-fullres-src"].textContent   # "https://graph.microsoft.com/v1.0/users('a6e04d01-7cbf-43bd-ab02-4a594039ff4c')/onenote/resources/1-a5a151f71313473aa67a9272e5088ee4!1-2881072d-ef79-44a1-ae15-8b8e0bcbe8b4/$value"
                $mediatype = $img.attributes["data-fullres-src-type"].textContent   # "image/png"
                if($mediatype -eq "application/octet-stream") {
                    $emfSrc = $fullResSrc
                    $fullResSrc = $img.attributes["src"].textContent
                    $mediatype = $img.attributes["data-src-type"].textContent
                }
            } catch { }
            if($null -ne $fullResSrc -or $null -ne $mediatype) {
                $parts = $fullResSrc -split "/"
                $filename = $parts[$parts.Count - 2]
                $ext = $mediatype.Replace("image/", "").Replace("application/", "")
                $out = "$exportpath\pages\$id\$filename.$ext"
                if((Test-Path $out) -eq $false) {
                    Write-Host "GET $($filename).$($ext)"
                    Invoke-PnPGraphMethod -Url $fullResSrc -OutFile $out
                    Start-Sleep -Milliseconds $requestPauseInMilliSeconds
                }
                # Update src in parsed html img tags to point to local
                $imgSrc = $img.attributes["src"].textContent
                $content = $content.Replace($imgSrc, "$filename.$ext")

                if($null -ne $emfSrc) {
                    $parts = $emfSrc -split "/"
                    $filename = $parts[$parts.Count - 2]
                    $out = "$exportpath\pages\$id\$filename.emf"
                    if((Test-Path $out) -eq $false) {
                        Write-Host "GET $($filename).emf"
                        Invoke-PnPGraphMethod -Url $emfSrc -OutFile $out
                        Start-Sleep -Milliseconds $requestPauseInMilliSeconds
                    }
                }
            }
        }
        $objTags = $html.all.tags("OBJECT")
        foreach($obj in $objTags) {
            # Download object file
            $objSrc = $null
            $filename = $null
            try {
                $objSrc = $obj.attributes["data"].textContent
                $filename = $obj.attributes["data-attachment"].textContent  # "https://graph.microsoft.com/v1.0/users('a6e04d01-7cbf-43bd-ab02-4a594039ff4c')/onenote/resources/1-682b90d53f5c00ea092565165223150d!1-3c32c77b-b10a-4e37-91a5-acd65f0b604b/$value"
            } catch {}
            if($null -ne $objSrc -or $null -ne $filename) {
                $out = "$exportpath\pages\$id\$filename"
                if((Test-Path $out) -eq $false) {
                    Write-Host "GET $filename"
                    Invoke-PnPGraphMethod -Url $objSrc -OutFile $out
                    Start-Sleep -Milliseconds $requestPauseInMilliSeconds
                }
                
                # replace <object> tags with file icon
                # <object data-attachment="woodchuck.docx" type="application/vnd.openxmlformats-officedocument.wordprocessingml.document" data="https://graph.microsoft.com/v1.0/users('a6e04d01-7cbf-43bd-ab02-4a594039ff4c')/onenote/resources/1-fab8ee7727980ca50dc091f02787b34e!1-3c32c77b-b10a-4e37-91a5-acd65f0b604b/$value" />
                $ext =  [System.IO.Path]::GetExtension($filename).Replace(".","").ToUpper()
                $col = "black";
                if($ext -eq "DOCX" -or $ext -eq "DOC") { $col = "#1554b7"; }
                if($ext -eq "XLSX" -or $ext -eq "XLS") { $col = "#10773e"; }
                if($ext -eq "PPTX" -or $ext -eq "PPT") { $col = "#c13c1c"; }
                if($ext -eq "PDF") { $col = "#B30B00"; }
                $fileIcon = $fileIconRaw.Replace("{{FILENAME}}", $filename).Replace("{{FILEEXT}}", $ext).Replace("{{EXTCOLOR}}", $col)
                $content = $content -replace "<object.*$([Regex]::Escape($filename)).*/>", $fileIcon
            }
        }

        # add a title since it is not included in html content
        $titleHTML = $titleHTMLRaw.Replace("{{PAGETITLE}}", $title).Replace("{{PAGEDATE}}", $page.lastModifiedDateTime.ToString("yyyy-MM-dd HH:mm"))
        $content = $content -replace "(<body.*>)", "`$1 `n $titleHTML"
        if($content -match "data-tag") {
            $content = $content -replace "(<head.*>)", "`$1 `n $tagStyling"
        }

        $content | Set-Content -Path $fileParsed -Encoding utf8 -Force
    }
    Write-Progress -Completed
}

function Get-PagesWithUnsupportedFeatures {
    $pageIds = Get-ChildItem -Path "$exportpath\pages" -Directory | select-object -ExpandProperty  "Name"
   
    $found = @()

    $count = 0
    ForEach($id in $pageIds) {
        $page = Get-Content "$exportpath\pages\$id\$id.html" -raw -ErrorAction SilentlyContinue
        $count++
        Write-Progress -Activity "Searching" -PercentComplete (Get-Percent $count $pageIds.Count)
        if($null -ne $page) {
            if($page.Contains("InkNode is not supported") -eq $true -or $page.Contains("IFrameNode is not supported") -eq $true) {
                $json = Get-Content "$exportpath\pages\$id\$id.json" -raw | ConvertFrom-Json
                $found += $json.title + " ($id)"
            }
        }
    }
    
    $found

    Write-Progress -Completed
}

function Export-PageStructure {

    write-host "Export sectionGroups.json..."
    $sectionGroups = Get-ChildItem -Path "$exportpath\sectionGroups\*.json" | ForEach-Object { (Get-Content -Raw $_.FullName | ConvertFrom-Json) }
    $count = 0
    $data = @()
    ForEach($group in $sectionGroups) {
        $count++
        Write-Progress -Activity "Processing section groups" -PercentComplete (Get-Percent $count $sectionGroups.Count)

        # Determine the path of the group.
        $path = @()
        $idPath = @()
        $g = $group.parentSectionGroup
        while($null -ne $g) {
            $sg = $sectionGroups | where { $_.id -eq $g.id } 
            $path += $sg.displayName
            $idPath += $sg.id
            $g = $sg.parentSectionGroup
        }

        ([array]::Reverse($path))
        ([array]::Reverse($idPath))

        $data += New-Object PSObject -property $([ordered]@{ 
            "Name" = $group.displayName
            "Id" = $group.id
            "Created" = $group.createdDateTime
            "Modified" = $group.lastModifiedDateTime
            "ParentSectionGroup" = $group.parentSectionGroup.displayName
            "ParentSectionGroupId" = $group.parentSectionGroup.id
            "Path" = $path -join "/"
            "IdPath" = $idPath -join "/"
        })
    }
    $data | ConvertTo-Json | Set-Content -Path "$exportpath\sectionGroups.json" -Encoding utf8NoBOM
    Write-Progress -Completed


    write-host "Export pages.json..."
    $sections = Get-ChildItem -Path "$exportpath\sections\*.json" | ForEach-Object { (Get-Content -Raw $_.FullName | ConvertFrom-Json) }
    $count = 0
    $data = @()
    ForEach($section in $sections) {
        $count++
        Write-Progress -Activity "Processing sections" -PercentComplete (Get-Percent $count $sections.Count)

        # Determine the path of the page. Pages are located in sections. Sections can be located in section groups. Section groups can also be located in section groups.
        $path = @()
        $idPath = @()
        $g = $section.parentSectionGroup
        while($null -ne $g) {
            $group = $sectionGroups | where { $_.id -eq $g.id } 
            $path += $group.displayName
            $idPath += $group.id
            $g = $group.parentSectionGroup
        }

        ([array]::Reverse($path))
        ([array]::Reverse($idPath))

        $data += New-Object PSObject -property $([ordered]@{ 
            "Name" = $section.displayName
            "Id" = $section.id
            "Created" = $section.createdDateTime
            "Modified" = $section.lastModifiedDateTime
            "ParentSectionGroup" = $section.parentSectionGroup.displayName
            "ParentSectionGroupId" = $section.parentSectionGroup.id
            "Path" = $path -join "/"
            "IdPath" = $idPath -join "/"
        })
    }
    $data | ConvertTo-Json | Set-Content -Path "$exportpath\sections.json" -Encoding utf8NoBOM
    Write-Progress -Completed

    write-host "Export pages.json..."
    $pageIds = Get-ChildItem -Path "$exportpath\pages" -Directory | select-object -ExpandProperty  "Name"
    $count = 0
    $data = @()
    ForEach($id in $pageIds) {
        $count++
        Write-Progress -Activity "Processing pages" -PercentComplete (Get-Percent $count $pageIds.Count)

        $page = Get-Content "$exportpath\pages\$id\$id.json" -raw | ConvertFrom-Json

        # Determine the path of the page. Pages are located in sections. Sections can be located in section groups. Section groups can also be located in section groups.
        $pageSection = $sections | where { $_.id -eq $page.parentSection.id } 
        $pagePath = @($pageSection.displayName)
        $pageIdPath = @($pageSection.id)
        $g = $pageSection.parentSectionGroup
        while($null -ne $g) {
            $group = $sectionGroups | where { $_.id -eq $g.id } 
            $pagePath += $group.displayName
            $pageIdPath += $group.id
            $g = $group.parentSectionGroup
        }

        ([array]::Reverse($pagePath))
        ([array]::Reverse($pageIdPath))

        $data += New-Object PSObject -property $([ordered]@{ 
            "Title" = ($page.title -eq "" ? "Untitled Page" : $page.title)
            "Id" = $page.id
            "Created" = $page.createdDateTime
            "Modified" = $page.lastModifiedDateTime
            "ParentSectionName" = $page.parentSection.displayName
            "ParentSectionId" = $page.parentSection.id
            "Order" = $page.order
            "Level" = $page.level
            "Path" = $pagePath -join "/"
            "IdPath" = $pageIdPath -join "/"
        })
    }
    Write-Progress -Completed
    $data | ConvertTo-Json | Set-Content -Path "$exportpath\pages.json" -Encoding utf8NoBOM
}

function Export-HTMLPreview {

    $notebook = Get-Content "$exportpath\notebook.json" -Raw | ConvertFrom-Json
    $pages = Get-Content "$exportpath\pages.json" -Raw | ConvertFrom-Json

    $data = @()
    ForEach($page in $pages) {
        $data += "<li>" + [System.Web.HttpUtility]::HtmlEncode($page.Path) + "/<a target='page' href='pages/" + $page.id + "/" + $page.id + ".PARSED.html'>" + [System.Web.HttpUtility]::HtmlEncode($page.Title) + "</a></li>`n"
    }
    $data = $data | Sort-Object
    $data = $data -join ""

    $previewHTML = '<html>
                        <head>
                            <title>{{TITLE}}</title>
                        </head>
                        <body>
                            <div style="display:flex; gap:10px; height: 100%">
                                <div style="max-width: 50%; resize: horizontal; overflow: auto;">
                                    <ul>
                                        {{DATA}}
                                    </ul>
                                </div>
                                <div style="display: flex; flex-grow: 1;">
                                    <iframe name="page" style="width: 100%;"></iframe>
                                </div>
                            </div>
                        </body>
                    </html>'

    $previewHTML = $previewHTML.Replace("{{TITLE}}", $notebook.displayName).Replace("{{DATA}}", $data)
    $previewHTML | Set-Content "$exportpath\preview.html" -Encoding utf8 -Force

}

# ------------------------------------------------------------------------------------------------------------

Write-Host "Connecting..." -ForegroundColor Cyan
connect-pnponline -Url $site -Interactive -ClientId $appId

if($stages -contains 1) {
    Write-Host "`n[[ STAGE 1 ]]" -ForegroundColor Yellow
    Write-Host "Get notebook..." -ForegroundColor Cyan
    $root = Invoke-PnPGraphMethod -Url $notebookUrl -All
    $root | ConvertTo-Json -Depth 100 | Set-Content -Path "$exportpath\notebook.json" -Encoding utf8 -Force
    Write-Host "`nGet section groups..." -ForegroundColor Cyan
    $sectionGroups = Get-SectionGroupsRecursive($root)
    if($detectRemovedItems -eq $true) { 
        Write-Host "`nPurge section groups..." -ForegroundColor Cyan
        Remove-DeletedSectionGroups $sectionGroups
    }
    Write-Host "`nGet sections in each group..." -ForegroundColor Cyan
    Get-Sections (@($root) + $sectionGroups) # treat root as a section group
} else {
    Write-Host "`nSkipping stage 1..." -ForegroundColor DarkGray
}

if($stages -contains 2) {
    Write-Host "`n[[ STAGE 2 ]]" -ForegroundColor Yellow
    Write-Host "Get page JSON data for unprocessed sections..." -ForegroundColor Cyan
    Get-PageJSON
} else {
    Write-Host "`nSkipping stage 2..." -ForegroundColor DarkGray
}

if($stages -contains 3) {
    Write-Host "`n[[ STAGE 3 ]]" -ForegroundColor Yellow
    Write-Host "`nGet page HTML..." -ForegroundColor Cyan
    Get-PageHTML
} else {
    Write-Host "`nSkipping stage 3..." -ForegroundColor DarkGray
}

if($stages -contains 4) {
    Write-Host "`n[[ STAGE 4 ]]" -ForegroundColor Yellow
    Write-Host "Download page resource files..." -ForegroundColor Cyan
    Get-PageLinkedResources
} else {
    Write-Host "`nSkipping stage 4..." -ForegroundColor DarkGray
}

if($stages -contains 5) {
    Write-Host "`n[[ STAGE 5 ]]" -ForegroundColor Yellow
    Write-Host "Listing pages missing content (ink drawings, loop components)..." -ForegroundColor Cyan
    Get-PagesWithUnsupportedFeatures
    Write-Host "`nCreate JSON exports..." -ForegroundColor Cyan 
    Export-PageStructure
    Write-Host "`nCreate preview html page..." -ForegroundColor Cyan 
    Export-HTMLPreview
} else {
    Write-Host "`nSkipping stage 5..." -ForegroundColor DarkGray
}

Write-Host "All done!" -ForegroundColor Green
