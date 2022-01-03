import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

import 'util.dart';
import 'helpers.dart';

class DocxMailMerge {
  ///Stores the original document file
  final List<int> docx;

  //TODO: add a preprocess fail state to fail softly
  bool _preprocessed = false;

  ///The potential archive that the docx file extracts to
  late Archive _archive;

  ///Stores the documents that merge fields could exist in
  Map<String, XmlDocument> _documents = {};

  //Stores the merge field nodes from the documents that can be replaced
  List<NodeField> _mergeFields = [];

  ///Create a DocxMailMerge with the raw file given
  DocxMailMerge(this.docx);

  ///Create a DocxMailMerge and [decode] the raw file given
  DocxMailMerge.preprocess(this.docx) {
    //TODO: This duplicates the work needed to be done for the second part of the preprocess stage for merge since merge will force it to run again. It would be good to have something like a _dirty flag to avoid duplicate work
    preprocess();
  }

  ///Preprocess the docx file given and identifies all the merge fields
  ///
  ///This sets-up the merge calls and that occur later
  ///
  ///[force] Recreates the [documents] and [mergeFields]. These are modified when a merge is ran
  void preprocess({bool force = false}) {
    if (!_preprocessed) {
      //extract all the files in the archive
      _archive = ZipDecoder().decodeBytes(docx);
    }

    if (!_preprocessed || force) {
      //get the documents that could have merge fields
      _documents = extractParts(_archive);
      //Get all merge fields in the documents and add them to the list
      _mergeFields = [];
      for (final entry in _documents.entries) {
        _mergeFields.addAll(getNodeFields(entry.value));
      }
    }

    _preprocessed = true;
  }

  ///Returns the merge field names for a docx file
  Set<String> get mergeFieldNames {
    preprocess();
    return _mergeFields.map((node) => node.field).toSet();
  }

  ///Does a merge on the fields defined in the [merge] keys with [merge]'s values
  ///
  ///[noProof] Have no proofchecking on the merged fields. Enabled by default as it reflects word's merge behavior
  ///
  ///[removeEmpty] Merge fields with no matching keys in [merge] are removed.
  ///
  ///This only modifies the in-memory representation of the document
  ///
  ///The in-memory will be set to the original document when the merge is called to clear any past merge operations
  List<int> merge(Map<String, String> merge, {bool noProof = true, bool removeEmpty = true}) {
    preprocess(force: true);
    mergeNodeFields(_mergeFields, merge, noProof: noProof, removeEmpty: removeEmpty);
    return mergeFiles(_archive, _documents);
  }
}
