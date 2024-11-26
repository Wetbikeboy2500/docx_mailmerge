import 'package:docx_mailmerge/docx_mailmerge.dart';
import 'package:xml/xml.dart';

///Represents all nodes that create a merge field
class NodeField {
  final List<XmlNode> elements;
  final String field;

  const NodeField(this.elements, this.field);
}

///Represent consecutive nodes that have merged together
class NodeFields {
  final List<XmlNode> elements;
  final List<String> fields;

  const NodeFields(this.elements, this.fields);
}

///Prints message if package is set to verbose
void log(Object? value) {
  if (DocxMailMerge.verbose) {
    // ignore: avoid_print
    print(value);
  }
}
