# Document Parser

## Prerequisites

[Visual Studio 2019 Community](https://visualstudio.microsoft.com/vs/community/)

## Setup

Build Solution file **DSnA.WebJob.DocumentParser.sln** with Visual Studio. This will generate the executable file called: *DSnA.WebJob.DocumentParser.exe*

## Configuration

App.config - add the key values for: 

    StorageConnectionString (only for blob storage mode): add storage connection string information in order to get access to the blob storage containers.

    OutputFileFormat: "csv" or "json" output file format

    StorageType: "blob" or "localstorage" values

## How it works?

The document parser tool allows you to extract the content of a document file (pdf, word) and create an output csv or json file with the document data classified into text, paragraphs, headings, sections, clauses, heading clauses and additional information.

To run this, you have two options:

**Azure blob storage**: upload the documents to be processed to the input blob storage and open a command prompt window located where the *DSnA.WebJob.DocumentParser.exe* file is located and then run it with the required arguments as below:

>DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3

Options:

>arg1: **Required** - blob input container name

>arg2: **Required** - blob virtual directory name/path (/ root level)

>arg3: Optional - file name filter (if not present, all documents within source folder will be processed)

The output files will be located in the blob storage *docparseoutput* container.

**Local storage**: upload the documents to the local folder in your file system to then open a command prompt window located where the *DSnA.WebJob.DocumentParser.exe* file is located and then run it with the required arguments as below:

>DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3

Options:

>arg1: **Required** - local storage source folder path

>arg2: **Required** - local storage output folder path

>arg3: Optional - file name filter (if not present, all documents within source folder will be processed)

The output files will be located in the output local folder.

