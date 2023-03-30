# Mail Merge to Separate Files Macro for Microsoft Word

This Microsoft Word macro automates the process of creating separate files for each record in a mail merge. It performs a mail merge using data from an Excel file and saves each merged document as a separate file in a specified folder, using a value from a specified field as part of the file name.

## Features

- Iterates through all records in the mail merge data source (Excel file)
- Performs a mail merge for each record and creates a new Word document with the merged data
- Saves the new Word document using the value from a specified field (e.g., Property Number) as part of the file name
- Closes the merged document and moves on to the next record in the data source

## Prerequisites

- Microsoft Word with the macro installed and added to the Ribbon
- A mail merge main document set up with merge fields and linked to an Excel data source

## Installation

1. Create the macro in the Visual Basic for Applications (VBA) editor in Microsoft Word.
2. Save the macro as a Word Add-In file (.dotm) in the Word Add-Ins folder.
3. Customize the Ribbon UI in Microsoft Word to add the macro button to a custom group.

## Usage

1. Open the mail merge main document in Microsoft Word.
2. Set the mail merge data source by linking to an Excel file containing the data.
3. Insert the appropriate merge fields into the main document using the data source
