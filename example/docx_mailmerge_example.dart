import 'dart:io';

import 'package:docx_mailmerge/docx_mailmerge.dart';

void main() {
  final merge = DocxMailMerge(File('test/files/original1.docx').readAsBytesSync());
  File('test/tmp/example.docx').writeAsBytesSync(merge.merge({'First_Name': 'hello world'}, removeEmpty: false));
}
