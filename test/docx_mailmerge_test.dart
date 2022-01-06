// ignore_for_file: avoid_print

import 'package:docx_mailmerge/docx_mailmerge.dart';

import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:docx_mailmerge/src/constants.dart';
import 'package:docx_mailmerge/src/helpers.dart';
import 'package:docx_mailmerge/src/util.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

///Merge fields
///For original1
const mergeFields1 = {
  'First_Name': '',
  'Last_Name': '',
  'Title': '',
  'City': '',
  'State': '',
  'Company_Name': '',
};

///Merge fields
///For original2
const mergeFields2 = {
  'First_Name': 'test1',
  'Last_Name': 'test2',
  'Title': 'test',
  'City': 'test4',
  'State': 'test5',
  'Company_Name': 'test3',
};

void main() {
  setUp(() {
    DocxMailMerge.verbose = true;
  });

  group('util', () {
    test('log', () {
      log('Testing');
    });
  });

  group('Decompress and compress without any changes', () {
    test('1', () {
      final file = File('test/files/original1.docx');
      final content = file.readAsBytesSync();
      print('read file');
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge({}, removeEmpty: false);
      print('merged');
      //get into a workable format
      final original = ZipDecoder().decodeBytes(content);
      final newOutput = ZipDecoder().decodeBytes(output);
      expect(original.equals(newOutput), isTrue);
    });

    test('2', () {
      final file = File('test/files/original2.docx');
      final content = file.readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge({}, removeEmpty: false);
      //get into a workable format
      final original = ZipDecoder().decodeBytes(content);
      final newOutput = ZipDecoder().decodeBytes(output);
      expect(original.equals(newOutput), isTrue);
    });
  });

  group('Get fields', () {
    test('1', () {
      final file = File('test/files/original1.docx');
      final content = file.readAsBytesSync();
      final m = DocxMailMerge(content);
      expect(mergeFields1.keys.toSet(), equals(m.mergeFieldNames));
    });

    test('2', () {
      final file = File('test/files/original2.docx');
      final content = file.readAsBytesSync();
      final m = DocxMailMerge.preprocess(content);
      expect(mergeFields2.keys.toSet(), equals(m.mergeFieldNames));
    });

    test('3', () {
      final file = File('test/files/original3.docx');
      final content = file.readAsBytesSync();
      final m = DocxMailMerge.preprocess(content);
      expect({'Last Name'}, equals(m.mergeFieldNames));
    });
  });

  //NOTE: These always fail due to differents/lack-of ids for elements. Looking at the diff for the xml does show that the merge fields do reflect how the merge should be like
  group('Merge file', () {
    test('1', () {
      final content = File('test/files/original1.docx').readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge(mergeFields1);
      //get into a workable format
      final merged = ZipDecoder().decodeBytes(File('test/files/original1_full_merge.docx').readAsBytesSync());
      final newOutput = ZipDecoder().decodeBytes(output);
      File('test/tmp/output1.docx').writeAsBytesSync(output);
      expect(merged.equals(newOutput), isTrue);
    },
        skip:
            'There are random trackers/markers that cannot be replicated to create an exact outcome of a regular mail merge');

    test('2', () {
      final content = File('test/files/original2.docx').readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge(mergeFields2);
      //get into a workable format
      final merged = ZipDecoder().decodeBytes(File('test/files/original2_full_merge.docx').readAsBytesSync());
      final newOutput = ZipDecoder().decodeBytes(output);
      File('test/tmp/output2.docx').writeAsBytesSync(output);
      final files = merged.files.where((element) => element.name.endsWith('.xml'));
      for (final file in files) {
        final alt = newOutput.findFile(file.name);
        if (alt != null && file.name != '[Content_Types].xml') {
          expect(XmlDocument.parse(Utf8Decoder().convert(alt.content)).toXmlString(pretty: true),
              equals(XmlDocument.parse(Utf8Decoder().convert(file.content)).toXmlString(pretty: true)));
        }
      }
      expect(merged.equals(newOutput), isTrue);
    },
        skip:
            'There are random trackers/markers that cannot be replicated to create an exact outcome of a regular mail merge');

    ///Tests against previous merge output
    ///
    ///This ensures that no changes occur without changing the test files themselves
    test('1 Previous output', () {
      final content = File('test/files/original1.docx').readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge(mergeFields1);
      //get into a workable format
      final merged = ZipDecoder().decodeBytes(File('test/files/original1_output.docx').readAsBytesSync());
      final newOutput = ZipDecoder().decodeBytes(output);
      expect(merged.equals(newOutput), isTrue);
    });

    ///Tests against previous merge output
    ///
    ///This ensures that no changes occur without changing the test files themselves
    test('2 Previous output', () {
      final content = File('test/files/original2.docx').readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge(mergeFields2);
      //get into a workable format
      final merged = ZipDecoder().decodeBytes(File('test/files/original2_output.docx').readAsBytesSync());
      final newOutput = ZipDecoder().decodeBytes(output);
      expect(merged.equals(newOutput), isTrue);
    });

    test('1 Nothing', () {
      final content = File('test/files/original1.docx').readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge({}, removeEmpty: false);
      //get into a workable format
      final merged = ZipDecoder().decodeBytes(content);
      final newOutput = ZipDecoder().decodeBytes(output);
      expect(merged.equals(newOutput), isTrue);
    });

    test('2 Nothing', () {
      final content = File('test/files/original2.docx').readAsBytesSync();
      //run merge with no changes
      final m = DocxMailMerge(content);
      final output = m.merge({}, removeEmpty: false);
      //get into a workable format
      final merged = ZipDecoder().decodeBytes(content);
      final newOutput = ZipDecoder().decodeBytes(output);
      expect(merged.equals(newOutput), isTrue);
    });
  });

  group('Helpers', () {
    test('Merge node fields empty', () {
      final builder = XmlBuilder();
      builder.element('r', nest: () {
        builder.element('t');
      });

      final f = builder.buildFragment();

      final parent = f.childElements.elementAt(0);
      final child = parent.childElements.elementAt(0);

      mergeNodeFields([
        NodeField([child], 'test')
      ], {}, removeEmpty: true);

      expect(parent.childElements, isEmpty);
    });

    ///For all the issues that can occur with a complex field
    group('Imcomplete Complex Field', () {
      test('parent', () {
        final builder = XmlBuilder();

        builder.element('fldChar', attributes: {'w:fldCharType': 'begin'}, namespace: NS.w, nest: () {
          builder.namespace(NS.w, 'w');
          builder.namespace(NS.ct, 'ct');
          builder.namespace(NS.mc, 'mc');
        });

        final f = builder.buildDocument();
        expect(getComplexFields(f), isEmpty);
      });

      test('sibling', () {
        final builder = XmlBuilder();

        builder.element('document', nest: () {
          builder.namespace(NS.w, 'w');
          builder.namespace(NS.ct, 'ct');
          builder.namespace(NS.mc, 'mc');
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('fldChar', attributes: {'w:fldCharType': 'begin'}, namespace: NS.w);
          });
        });

        final f = builder.buildDocument();
        expect(getComplexFields(f), isEmpty);
      });
      test('sibling no instr', () {
        final builder = XmlBuilder();

        builder.element('document', nest: () {
          builder.namespace(NS.w, 'w');
          builder.namespace(NS.ct, 'ct');
          builder.namespace(NS.mc, 'mc');
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('fldChar', attributes: {'w:fldCharType': 'begin'}, namespace: NS.w);
          });
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('instrText', namespace: NS.w);
          });
        });

        final f = builder.buildDocument();
        expect(getComplexFields(f), isEmpty);
      });
      test('sibling instr issue', () {
        final builder = XmlBuilder();

        builder.element('document', nest: () {
          builder.namespace(NS.w, 'w');
          builder.namespace(NS.ct, 'ct');
          builder.namespace(NS.mc, 'mc');
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('fldChar', attributes: {'w:fldCharType': 'begin'}, namespace: NS.w);
          });
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('instrText', attributes: {'w:instr': ''}, namespace: NS.w);
          });
        });

        final f = builder.buildDocument();
        expect(getComplexFields(f), isEmpty);
      });
      test('no separate', () {
        final builder = XmlBuilder();

        builder.element('document', nest: () {
          builder.namespace(NS.w, 'w');
          builder.namespace(NS.ct, 'ct');
          builder.namespace(NS.mc, 'mc');
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('fldChar', attributes: {'w:fldCharType': 'begin'}, namespace: NS.w);
          });
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('instrText', namespace: NS.w, nest: () {
              builder.text('MERGEFIELD "name"');
            });
          });
        });

        final f = builder.buildDocument();
        expect(getComplexFields(f), isEmpty);
      });
      test('no end', () {
        final builder = XmlBuilder();

        builder.element('document', nest: () {
          builder.namespace(NS.w, 'w');
          builder.namespace(NS.ct, 'ct');
          builder.namespace(NS.mc, 'mc');
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('fldChar', attributes: {'w:fldCharType': 'begin'}, namespace: NS.w);
          });
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('instrText', namespace: NS.w, nest: () {
              builder.text('MERGEFIELD "name"');
            });
          });
          builder.element('r', namespace: NS.w, nest: () {
            builder.element('fldChar', attributes: {'w:fldCharType': 'separate'}, namespace: NS.w);
          });
        });

        final f = builder.buildDocument();
        expect(getComplexFields(f), isEmpty);
      });
    });
  });
}

extension ArchiveEquals on Archive {
  bool equals(Archive a) {
    return every((file) {
      String name = file.name;

      final other = a.findFile(name);

      if (other == null) {
        print('Not found');
        return false;
      }

      if (other.size != file.size) {
        print('Wrong size');
        return false;
      }

      if (other.content != null && file.content != null) {
        if (other.content is List && file.content is List) {
          if (other.content.length == file.content.length) {
            for (int i = 0; i < 0; i++) {
              if (other.content[i] != file.content[i]) {
                print('Wrong content');
                return false;
              }
            }
          } else {
            print('Wrong length');
            return false;
          }
        } else {
          print('Wrong types');
          return false;
        }
      } else {
        print('Null content');
        return false;
      }

      return true;
    });
  }
}
