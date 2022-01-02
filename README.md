Performs a mail merge operation on DOCX files. There are no platform dependant external dependencies which allows this package to run anywhere. Mail merge can also run in the browser!

## Features

- Merge values with merge fields within a DOCX document
- Get merge fields a document contains
- Evergy merge works on a fresh copy of the document
    - Enables batch merging
- Enable or disable proofreading on merged fields
    - Proofreading is disabled on merged fields by default to better match behavior seen in Word
- Enable or disable the removal of merge fields that do not have a corresponding key/value to merge with

## Getting started

This package reads files as bytes through the data structure of a `List<int>`. For web, this project has an example in the `/example/web` folder for your use (JS compiled separately)

## Usage

To setup a document to be merged with, pass a `List<int>` of the file to the `DocxMailMerge` constructor.

```dart
DocxMailMerge(File('test/files/original1.docx').readAsBytesSync())
```

There is also a named constructor which will unzip and read the merge fields immediately rather than on demand.

```dart
DocxMailMerge.preprocess(File('test/files/original1.docx').readAsBytesSync())
```



## Additional information

TODO: Tell users more about the package: where to find more information, how to
contribute to the package, how to file issues, what response they can expect
from the package authors, and more.
