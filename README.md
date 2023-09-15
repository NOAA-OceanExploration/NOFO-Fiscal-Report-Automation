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
- A Local Python Installation (3.9)

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

- **Generate All Figures**: Calls all figure generation functions (`figureOne` to `figureSix`).

#### 1. Generate Figure One

- Copies specific columns from an original spreadsheet.
- Sets headers and row names.
- Inserts data into a new sheet.
- Creates a column chart.

#### 2. Generate Figure Two

- Similar to Figure One but with different headers and data columns.
- Sorts data.
- Creates a column chart.

#### 3. Generate Figure Three

- Similar to Figure One but with different headers and data columns.
- Sorts data.
- Creates a column chart.

#### 4. Generate Figure Four

- Similar to Figure One but with different headers and data columns.
- Sorts data.
- Creates a column chart.

#### 5. Generate Figure Five

- Similar to Figure One but with different headers and data columns.
- Sorts data.
- Creates a column chart.

#### 6. Generate Figure Six

- Calculates sums of specific columns from the original sheet.
- Copies a specific column from the original sheet.
- Sorts data.
- Creates a column chart.

## Spreadsheet Clearing

- **Clear Spreadsheet**: 
  - Clears all content, formats, and comments from 'Sheet2'.
  - Removes all charts from 'Sheet2'.

## Suggested Workflow
1. **Populate NOFO Fiscal Report Figures**: Create a copy of the template document and use the functions available after installation to populate the sheet.
2. **Copy Figures into NOFO Report**: Create a copy of the NOFO report template document and use the functions to move the figures and then generate table.
3. **Use Python Script to Generate Node Graph**: Copy the python script to a local python runtime and include a copy of S&T Metrics in the working directory and copy the resulting node graph into the NOFO Report.

## Contributing

Contributions are welcome!

## License

This repository is a scientific product and is not official communication of the National Oceanic and Atmospheric Administration, or the United States Department of Commerce. All NOAA GitHub project code is provided on an 'as is' basis and the user assumes responsibility for its use. Any claims against the Department of Commerce or Department of Commerce bureaus stemming from the use of this GitHub project will be governed by all applicable Federal law. Any reference to specific commercial products, processes, or services by service mark, trademark, manufacturer, or otherwise, does not constitute or imply their endorsement, recommendation or favoring by the Department of Commerce. The Department of Commerce seal and logo, or the seal and logo of a DOC bureau, shall not be used in any manner to imply endorsement of any commercial product or activity by DOC or the United States Government.

This project is licensed under the MIT License - see [LICENSE.md](LICENSE.md) for more information.
