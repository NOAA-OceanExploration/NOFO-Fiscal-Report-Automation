# NOFO Fiscal Report Automation

This project provides a Google Apps Script to automate the handling of fiscal report figures, graphs, tables, and other elements within a Google Document, using Google Sheets as the data source.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Functions](#functions)
- [Contributing](#contributing)
- [License](#license)

## Introduction

The provided code automates various tasks such as inserting charts at specific bookmarks, clearing images and charts, copying graphs from Google Sheets, and managing tables within a Google Document. This enables streamlined reporting, particularly for fiscal data.

## Features

- **Insert Charts at Specific Locations**: Allows insertion of charts from Google Sheets to Google Document at specified bookmarks.
- **Clear Images and Charts**: Provides the ability to clear all the images and charts from the Google Document.
- **Copy Graphs from a Spreadsheet**: Enables copying specific graphs from a target sheet to a Google Document.
- **Generate and Manage Tables**: Facilitates creating, formatting, and deleting tables within a Google Document.

## Prerequisites

- Google Account
- Google Document
- Google Sheet with charts and data to be used

## Installation of Google Doc Code (graph.gs)

1. **Open Google Document**: Open the NOFO Template Google Doc where you want to run the script. (https://docs.google.com/document/d/1pLWFshH96w_NLUpRYqDXIeCra5w96dM0WoKK8fvCyMs/edit?usp=sharing)
2. **Access Script Editor**: Click on `Extensions` > `Script Editor`.
3. **Copy the Code**: Paste the entire (javascript) code from this repository (graph.gs) into the script editor.
4. **Save and Reload**: Press Ctrl + S to save, and reload the Google Document.
5. **Grant Permissions**: Allow the script to access and manage your documents and spreadsheets.
6. **Target Correct Documenst**: Assign the correct document IDs to the variables targetSheet (the google sheet with the graphs for the report figures), stMetricsSheet (the S&T Metrics Sheet), docId (the document ID of the report itself), fiscalYear (whatever the current fiscal year is).

## Installation of NOFO Fiscal Report Figures (populate_spreadsheet.gs)
1. **Open Google Sheet**: Open the NOFO Fiscal Report Figures where you want to run the script which is in the NOFO automation folder.
2. **Access Script Editor**: Click on `Extensions` > `Script Editor`.
3. **Copy the Code**: Paste the entire (javascript) code from this repository (populate_spreadsheet.gs) into the script editor.
4. **Save and Reload**: Press Ctrl + S to save, and reload the Google Document.
5. **Target Correct Documenst**: Assign the correct document IDs to the variables originalSpreadsheet (S&T Metrics Sheet), and newSpreadsheet (Current Spreadsheet).

## Usage

### Menu for Google Doc

After the installation, a new menu item 'Functions' will appear in the Google Document toolbar. It includes:

- **Copy Graphs From Spreadsheet**: Inserts graphs from specified Google Sheet.
- **Clear Images And Charts**: Removes images and charts.
- **Generate Table**: Creates a table from data in a Google Sheet.
- **Remove Table**: Deletes the generated table.

### Menu for Google Spreadsheet

After the installation, a new menu item 'Functions' will appear in the Google Document toolbar. It includes:



### Document Setup

Ensure that the document and spreadsheets are set up with the correct IDs and bookmarks. You may need to update the `targetSheet`, `stMetricsSheet`, and `docId` variables as per your configuration.

## Functions

Below are detailed descriptions of the main functions:

### `insertChartAtBookmark(targetSheetID, targetBookmarkId, chart_pos)`

Inserts a specific chart from the Google Sheet into the Google Document at a given bookmark.

### `clearImagesAndCharts()`

Clears all the images and charts from the Google Document.

### `copyGraphsFromSpreadsheet()`

Copies multiple figures from a target sheet to bookmarks within the Google Document.

### `copyToTable()`

Copies data from a Google Sheet into a formatted table within a Google Document, applying specific filters and styling.

### `deleteTable()`

Deletes a table at a specific bookmark within the Google Document.

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## License

This project is licensed under the MIT License - see [LICENSE.md](LICENSE.md) for more information.
