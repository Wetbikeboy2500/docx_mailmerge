// ignore_for_file: avoid_print

import 'dart:io';

void main(List<String> args) {
  if (args.contains('clean')) {
    clean();
    return;
  }

  final Iterable<String> paths =
      Directory('lib').listSync(recursive: true).map((e) => e.path.replaceAll('\\', '/')).where((element) {
    if (!element.endsWith('.dart') ||
        element.endsWith('lib/docx_mailmerge.dart')) {
      return false;
    }

    return true;
  }).map((e) => "import 'package:docx_mailmerge/${e.replaceFirst('lib/', '')}';");
  final File testFile = File('test/docx_mailmerge_test.dart');
  final List<String> lines = testFile.readAsLinesSync();

  lines.removeWhere((line) => line.startsWith('import \'package:docx_mailmerge'));

  lines.insertAll(0, paths);

  Directory('coverage').createSync();

  File('test/tmp.dart').writeAsStringSync(lines.join('\r\n'));

  Process.runSync('dart', ['test', '--file-reporter', 'json:reports/tests.json', 'test/tmp.dart', '--coverage=.']);
  Process.runSync('dart', [
    'run',
    'coverage:format_coverage',
    '--packages=.packages',
    '-i',
    './test/tmp.dart.vm.json',
    '-l',
    '-o',
    './coverage/lcov.info',
    '--report-on=lib'
  ]);

  clean();
}

void clean() {
  final List<File> files = [
    File('test/tmp.dart'),
    File('test/tmp.dart.vm.json'),
  ];

  for (var file in files) {
    if (file.existsSync()) {
      file.deleteSync();
    }
  }
}