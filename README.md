# VBA Text File Import Automation

## Project Overview

This project demonstrates how Microsoft Excel VBA can be used to automate the process of importing text files into a workbook. Instead of manually locating, opening, and copying data from external text files, the user can select a file directly from the system and import the data with minimal effort.

The solution is designed to improve efficiency, reduce manual work, and provide a user-friendly experience for handling external data sources.

## Features

- Import text files directly into Excel using VBA.
- Browse and select files from the system using `Application.GetOpenFilename`.
- Retrieve files from a designated **Data** folder.
- Automatically open selected text files.
- Copy and paste imported data into the target worksheet.
- Improve performance using `Application.ScreenUpdating = False`.
- Reduce user interaction and repetitive tasks through automation.
- Error handling for invalid file selections or cancellations.

## VBA Concepts Used

### Application.GetOpenFilename
Allows users to browse and select a text file from their local system.

### Application Object
Used to interact with Excel application settings and manage workbook operations.

### Copy and Paste Automation
Automatically transfers data from the imported text file into the required worksheet.

### Screen Updating
Improves execution speed and provides a smoother user experience by temporarily disabling screen refreshes during the import process.

```vb
Application.ScreenUpdating = False
' VBA Code
Application.ScreenUpdating = True
```

### File Handling
Used to locate and access text files stored within the project’s Data folder.

