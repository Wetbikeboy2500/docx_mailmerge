import 'dart:io';

import 'package:docx_mailmerge/docx_mailmerge.dart';

void main() {
  //Read a docx file as bytes and pass it to the constructor
  final merge =
      DocxMailMerge(File('test/files/original1.docx').readAsBytesSync());
  //This is just to ensure the output directory exists
  Directory('test/tmp').createSync();
  //Writes the generated merge information(bytes) to the output file
  File('test/tmp/example.docx').writeAsBytesSync(
      merge.merge({'First_Name': 'hello world'}, removeEmpty: false));
}
