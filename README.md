# HTML Lab Viewer

An empty HTML file picker for PowerPoint. No teaching files or sample labs are included.

Choose a self-contained HTML file from your computer. The viewer reads it into an isolated frame in memory; it does not upload it, store it, or save it in PowerPoint. Reselect the file after reopening the add-in. Selected HTML can load external scripts or services: only open files you trust. Relative companion files are not automatically available.

## Mac setup
Download manifest.xml. Save your presentation and quit PowerPoint. Replace the previous local manifest with this file in:

`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`

If the old file has a different name, remove that old manifest to avoid duplicate add-in IDs. Reopen PowerPoint, open My Add-ins and insert HTML Lab Viewer. Existing inserted instances may need to be removed and reinserted.

No local server, Node.js or local certificate is required for this hosted version. Internet is required to load the viewer and Microsoft Office.js. Test file selection and interaction in PowerPoint Slide Show before classroom use.

The public GitHub site and source contain only this generic viewer, manifest and instructions. GitHub and Microsoft receive normal webpage/library requests. The viewer includes no analytics or file-upload endpoint.
