# LabelingPropertiesReport
A utility for generating a report about labeling properties in an ArcGIS map document.

Implemented as an Add-In in ArcGIS Desktop, see releases for binaries versions of the esriAddIn file.

1. Download the esriAddIn file from the latest version at https://github.com/williamscraigm/LabelingPropertiesReport/releases
1. Double click on the LabelingPropertiesReport.esriAddIn file to install it
1. Open ArcMap
1. Click Customize > Customize Mode
1. Click on the Commands tab
1. Browse to the "Add-In Controls" group
1. Drag and drop the "Create Labeling Properties Report" command onto a toolbar
1. Open the map document you wish to create a report for.
1. Click the tool
1. The tool will run and produce a log file in the directory LabelingReport in your profile's Documents folder.  The file name will be yourDocumentName.log
1. The tool displays a message box with the report path when it is done.
