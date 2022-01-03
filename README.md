[![pub package](https://img.shields.io/pub/v/docx_mailmerge.svg)](https://pub.dev/packages/docx_mailmerge)
[![Build](https://github.com/Wetbikeboy2500/docx_mailmerge/actions/workflows/build.yml/badge.svg)](https://github.com/Wetbikeboy2500/docx_mailmerge/actions/workflows/build.yml)
[![codecov](https://codecov.io/gh/Wetbikeboy2500/docx_mailmerge/branch/master/graph/badge.svg?token=0BRNYQPEJK)](https://codecov.io/gh/Wetbikeboy2500/docx_mailmerge)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Performs a mail merge operation on DOCX files. There are no platform dependant external dependencies which allows this package to run anywhere. Mail merge can even run in the browser!

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

The `mergeFieldNames` getter will return a `Set<String>` for the merge fields that exist in a document.

```dart
DocxMailMerge(File('test/files/original1.docx').readAsBytesSync()).mergeFieldNames
```

The merge operation takes a `Map<String, String>` for the key/values to be merged. There are also optional parameters which are documented in the code. The merge operation returns a `List<int>` which is the new file with the merged fields. Every merge operation reuses the original file and runs a merge on it allowing for repeat merge operations to be run on a single `DocxMailMerge` object without needing to create more.

```dart
DocxMailMerge(File('test/files/original1.docx').readAsBytesSync()).merge({'First_Name': 'hello world'}, removeEmpty: false)
```

## Additional information

This package comes from the issue of not having an easy cross-platform mailmerge package. This is inspired by the [docx-mailmerge](https://github.com/Bouke/docx-mailmerge) python package. The python package is more mature and may cover cases not yet addressed by this package, but this dart package does have an advantage. With the many different target platforms that dart can compile to, this package can be adopted easier into more ecosystems (Did i mention it can run in the browser?!).
