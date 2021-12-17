import 'package:xml/xml.dart';

extension StringStrip on String {
  ///Removes a pattern if it happens at the start and end of a string
  String strip(String str) {
    if (startsWith(str) && endsWith(str)) {
      return substring(str.length, length - str.length);
    }

    return this;
  }
}

///Represents all nodes that create a merge field
class NodeField {
  final List<XmlNode> elements;
  final String field;

  const NodeField(this.elements, this.field);
}

class NodeFields {
  final List<XmlNode> elements;
  final List<String> fields;

  const NodeFields(this.elements, this.fields);
}

