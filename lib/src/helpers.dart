//This file contains a lot of funational code. This allows for testing and abstracting the different processes involved with merge fields

import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:commandline_splitter2/commandline_splitter2.dart';
import 'package:docx_mailmerge/src/util.dart';
import 'package:xml/xml.dart';
import 'constants.dart';

///Decode an archive file to an xml document
XmlDocument decodeArchiveFile(ArchiveFile f) => _decodeRawXML(f.content);

///Decode a raw file input to an xml document
XmlDocument _decodeRawXML(List<int> content) {
  //decode as UTF8 so it can then be parsed as xml
  return XmlDocument.parse(Utf8Codec().decode(content));
}

///Extracts the files that a mail merge will run on
///
///Returns the path name along with the xml needed to be parsed
Map<String, XmlDocument> extractParts(Archive archive) {
  //Load in content types to identify the needed xml files to manipulate
  final ArchiveFile contentTypes = archive.files.firstWhere((file) => file.name == '[Content_Types].xml');

  //decode the main content types to get the file paths that can be used
  final XmlDocument doc = decodeArchiveFile(contentTypes);
  final Iterable<XmlElement> elements = doc.findAllElements('Override', namespace: NS.ct);
  //get the names/directories in the archive for the headers, body, and footer of the doc
  final Map<String, XmlDocument> parts = {};
  for (final element in elements) {
    //only check specific parts of the word doc
    if (contentTypeParts.contains(element.getAttribute('ContentType'))) {
      //get the part name
      String? partName = element.getAttribute('PartName');
      if (partName != null) {
        //remove a starting / to match archive file names
        if (partName.startsWith('/')) {
          partName = partName.substring(1);
        }

        parts[partName] = _decodeRawXML(archive.files.firstWhere((file) => partName == file.name).content);
      }
    }
  }

  return parts;
}

///Gets all fields, simple or complex, from the [document]
List<NodeField> getNodeFields(XmlDocument document) {
  List<NodeField> nodeFields = [];

  nodeFields.addAll(getSimpleFields(document));
  nodeFields.addAll(getComplexFields(document));

  return nodeFields;
}

///Gets only simple fields from [document]
List<NodeField> getSimpleFields(XmlDocument document) {
  List<NodeField> nodeFields = [];
  //get simple fields
  Iterable<XmlElement> fields = document.findAllElements('fldSimple', namespace: NS.w);
  //go through each element
  for (final field in fields) {
    //get specific attribute with the field name
    final String? instr = field.getAttribute('instr', namespace: NS.w);
    if (instr != null) {
      //split instr like a shell-like command
      Iterable<String> args = split(instr).map((e) => e.strip('"'));
      //check if merge field exists
      if (args.length > 1 && args.elementAt(0) == 'MERGEFIELD') {
        //add field name to the set
        nodeFields.add(NodeField([field], args.elementAt(1)));
      }
    }
  }
  return nodeFields;
}

///Gets only complex fields from [document]
List<NodeField> getComplexFields(XmlDocument document) {
  List<NodeField> nodeFields = [];

  Iterable<XmlElement> fields = document.findAllElements('fldChar', namespace: NS.w);

  //Get only the beginning nodes
  fields = fields.where((element) => element.getAttribute('fldCharType', namespace: NS.w) == 'begin');

  //get complex fields
  //http://officeopenxml.com/WPfields.php
  //complex
  //find <w:fldChar w:fldCharType="begin"/> (only element that is replaced)
  //Go to parent
  //find instr and get mergefield <w:instrText xml:space="preserve"> MERGEFIELD Last_Name </w:instrText>
  //find separate <w:r><w:fldChar w:fldCharType="separate"/>
  //next <w:r> children until end contain the mergefield value
  //get adjacent until <w:r> with <w:fldChar w:fldCharType="end"/>
  //save node
  for (final field in fields) {
    List<XmlNode> elements = [];

    //moves through "states" begin -> instr -> separate -> [nodes] -> end
    //begin text run
    XmlNode? currentNode = field.parent;
    if (currentNode == null) {
      print('No parent to begin');
      continue;
    }
    elements.add(currentNode);

    //instr text run
    currentNode = currentNode.nextSibling;
    if (currentNode == null) {
      print('No sibling to start');
      continue;
    }
    //get instr
    XmlElement? instruction = currentNode.getElement('instrText', namespace: NS.w);
    if (instruction == null) {
      print('Sibling did not have instr');
      continue;
    }
    elements.add(currentNode);
    //get the merge field
    Iterable<String> args = split(instruction.text.trim()).map((e) => e.strip('"'));
    if (args.isEmpty || args.length < 2 || args.elementAt(0) != 'MERGEFIELD') {
      print('Args or mergefield instr issue');
      continue;
    }
    //get field name
    String fieldName = args.elementAt(1);

    //separate
    currentNode = currentNode.nextSibling;
    if (currentNode == null ||
        currentNode.getElement('fldChar', namespace: NS.w)?.getAttribute('fldCharType', namespace: NS.w) !=
            'separate') {
      print('separate not found');
      continue;
    }
    elements.add(currentNode);

    //Go till end
    bool end = false;
    currentNode = currentNode.nextSibling;

    while (currentNode != null) {
      elements.add(currentNode);

      if (currentNode.getElement('fldChar', namespace: NS.w)?.getAttribute('fldCharType', namespace: NS.w) == 'end') {
        end = true;
        break;
      }

      currentNode = currentNode.nextSibling;
    }

    if (end) {
      nodeFields.add(NodeField(elements, fieldName));
    } else {
      print('Could not find complex field\'s end');
    }
  }

  return nodeFields;
}

List<int> mergeFiles(Archive archive, Map<String, XmlDocument> documents) {
  final Archive a = Archive();

  for (ArchiveFile file in archive.files) {
    //get document for archive file
    final XmlDocument? document = documents[file.name];
    if (document != null) {
      //create the xml string to be written
      final List<int> content = Utf8Codec().encode(document.toXmlString());
      //add the document to the archive
      a.addFile(ArchiveFile(file.name, content.length, content));
    } else {
      //add the current archive file to the new archive
      a.addFile(file);
    }
  }

  //return zip or nothing
  return ZipEncoder().encode(a) ?? [];
}

void mergeNodeFields(List<NodeField> nodes, Map<String, String> merge, {bool noProof = true, bool removeEmpty = true}) {

  //skip NodeField if not in merge to avoid any changes
  if (!removeEmpty) {
    nodes.removeWhere((element) => !merge.containsKey(element.field));
  }

  //merge NodeFields together so they are in the same text element
  List<NodeFields> nodeFields = nodes.map((e) => NodeFields(e.elements, [e.field])).toList();
  List<NodeFields> mergedNodes = [];
  while (nodeFields.isNotEmpty) {
    //keep merging sibling nodes
    if (nodeFields.length > 1 && nodeFields[1].elements.first == nodeFields.first.elements.last.nextSibling) {
      nodeFields[0] = NodeFields([...nodeFields.first.elements, ...nodeFields[1].elements],
          [...nodeFields.first.fields, ...nodeFields[1].fields]);
      nodeFields.removeAt(1);
    } else {
      //Add to merged nodes if not sibling or only one node is left
      mergedNodes.add(nodeFields.removeAt(0));
    }
  }

  for (final NodeFields node in mergedNodes) {
    if (node.elements.isNotEmpty) {
      XmlNode first = node.elements.first;
      XmlNode? parent = node.elements.first.parent;
      if (parent != null) {
        var builder = XmlBuilder();
        //add namespaces
        builder.namespace(NS.w, 'w');
        builder.namespace(NS.ct, 'ct');
        builder.namespace(NS.mc, 'mc');

        //build the text thing
        bool hasOne = false;
        builder.element('r', namespace: NS.w, nest: () {
          if (noProof) {
            builder.element('rPr', namespace: NS.w, nest: () {
              builder.element('noProof', namespace: NS.w);
            });
          }
          builder.element('t', namespace: NS.w, nest: () {
            for (final field in node.fields) {
              if (merge.containsKey(field)) {
                hasOne = true;
                builder.text(merge[field]!);
              }
            }
          });
        });

        //require one field to be merged with to make the changes
        if (hasOne) {
          //remove all other nodes except the first
          for (int i = 1; i < node.elements.length; i++) {
            parent.children.remove(node.elements[i]);
          }
          //replace for the first node
          first.replace(builder.buildFragment());
        } else if (removeEmpty) { 
          //remove everything
          for (int i = 0; i < node.elements.length; i++) {
            parent.children.remove(node.elements[i]);
          }
        }
      }
    }
  }
}
