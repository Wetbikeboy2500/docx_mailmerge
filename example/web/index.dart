import 'dart:convert';
import 'dart:html';
import 'package:docx_mailmerge/docx_mailmerge.dart';

void main() {
  final element = document.querySelector('#submit');
  element?.addEventListener('click', (event) {
    final input = document.querySelector('#file_input') as InputElement;
    if (input.files != null && input.files!.isNotEmpty) {
      final file = input.files!.first;
      final reader = FileReader();
      reader.readAsArrayBuffer(file);
      reader.onLoad.first.then((value) {
        final dmm = DocxMailMerge(reader.result as List<int>);
        //for original2.docx in test/files
        final result = dmm.merge({
          'First_Name': 'test1',
          'Last_Name': 'test2',
          'Title': 'test',
          'City': 'test4',
          'State': 'test5',
          'Company_Name': 'test3',
        });
        final content = base64Encode(result);
        AnchorElement(
            href:
                "data:application/octet-stream;charset=utf-16le;base64,$content")
          ..setAttribute("download", "merged.docx")
          ..click();
      });
    }
  });
}
