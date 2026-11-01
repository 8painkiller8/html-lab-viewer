# HTML Lab Viewer

A public, empty HTML viewer for PowerPoint. No sample labs or teaching files are hosted here.

## Using it
Choose a self-contained HTML file. Inside PowerPoint the viewer saves a copy of its source and your zoom preference through Office document settings. Wait for the save confirmation, then save the PowerPoint presentation (Command-S). On reopening, the saved source is restored. Anyone receiving the PPT receives the embedded HTML too. The original file is unchanged. Changes to that original require selecting it again. Simulation state inside the loaded page is not captured.

Three views, saved with the presentation:

- **Fill window** (default): the page is laid out at the chosen page width and stretched to fill the add-in box exactly, so the page sees a window the same shape as the box. Pages that adapt to their window height fit completely; taller pages scroll inside the box. Here + and − step the page width down or up (a narrower page looks larger).
- **Whole page**: the full page height is shrunk to fit the box. Tall pages can look small.
- **Manual zoom**: + and − after choosing Whole page set a fixed zoom with scrolling.

Page width can be 960 to 1920 pixels (default 1280). Hide controls gives more space.

Save in presentation retries saving. Clear file removes the stored HTML copy when saved; save the PPT afterward. Large HTML files may exceed Office settings capacity. The viewer conservatively limits embedding to about 1.5 MB of UTF-16 serialized data and displays an error rather than claiming success. Standalone browser use is preview-only.

The viewer does not upload selected HTML to GitHub or an application server. Only open trusted HTML: its own scripts may contact external services. Relative companion files are not automatically available. The isolated page cannot access the parent viewer or Office APIs.

## Mac installation
Put manifest.xml in `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`, replacing the old manifest with the same add-in ID. Restart PowerPoint and insert HTML Lab Viewer from My Add-ins.

No local server or local certificate is needed. Internet is required to load this viewer, Office.js and external resources used by your HTML. Verify save/reopen and interaction in your PowerPoint Slide Show before classroom use. Automated tests exercise settings with a simulated Office host, not the actual PowerPoint application.
