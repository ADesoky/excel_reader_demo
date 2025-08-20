import 'dart:io' show File; // على الويب بيتجاهلها تلقائيًا
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart' show kIsWeb, compute, Uint8List;
import 'package:flutter/material.dart';

/// صفحة لقراءة ملف Excel (.xlsx) وعرضه في DataTable.
/// - تستخدم FilePicker لاختيار الملف (لا تحتاج صلاحيات في معظم الحالات)
/// - تدعم الويب (bytes من picker) والموبايل (path أو bytes)
class ExcelReaderPage extends StatefulWidget {
  const ExcelReaderPage({super.key});

  @override
  State<ExcelReaderPage> createState() => _ExcelReaderPageState();
}

class _ExcelReaderPageState extends State<ExcelReaderPage> {
  /// مخزن البيانات: كل عنصر = صف، وكل صف = قائمة خلايا نصية.
  List<List<String>> _excelData = [];

  bool _loading = false;
  String? _fileName;
  String? _sheetName;

  /// يحوّل bytes → مصفوفة نصوص (تشغيلها داخل isolate لتحسين الأداء)
  static List<List<String>> _parseExcelBytes(Uint8List bytes) {
    final excel = Excel.decodeBytes(bytes);
    if (excel.tables.isEmpty) return [];
    final firstSheet = excel.tables.keys.first;
    final sheet = excel.tables[firstSheet];
    if (sheet == null) return [];

    String cellToString(dynamic v) {
      if (v == null) return "";
      if (v is DateTime) return v.toIso8601String();
      return v.toString();
    }

    final rows = sheet.rows;
    return rows
        .map((row) => row.map((cell) => cellToString(cell?.value)).toList())
        .toList();
  }

  /// يفتح FilePicker ويقرأ الملف ثم يحدّث الحالة.
  Future<void> _pickAndReadExcel() async {
    try {
      setState(() => _loading = true);

      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: const ['xlsx'],
        withData: true, // مهم للويب وللتخلّي عن القراءة من المسار
      );

      if (result == null) {
        setState(() => _loading = false);
        return;
      }

      _fileName = result.files.single.name;

      // bytes أولًا؛ لو مش متاحة استخدم path (موبايل غالبًا)
      final bytes = result.files.single.bytes ??
          await File(result.files.single.path!).readAsBytes();

      // parse داخل isolate (خصوصًا للملفات الكبيرة)
      final data = await compute(_parseExcelBytes, bytes);

      setState(() {
        _excelData = data;
        _sheetName = data.isNotEmpty ? "Sheet1" : null; // اسم افتراضي
        _loading = false;
      });
    } catch (e, st) {
      setState(() => _loading = false);
      debugPrint("❌ Excel read error: $e\n$st");
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('فشل قراءة الملف: $e')),
        );
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final hasData = _excelData.isNotEmpty;

    return Scaffold(
      appBar: AppBar(title: const Text("Excel Reader")),
      body: Column(
        children: [
          Padding(
            padding: const EdgeInsets.all(12),
            child: Row(
              children: [
                ElevatedButton.icon(
                  onPressed: _loading ? null : _pickAndReadExcel,
                  icon: const Icon(Icons.upload_file),
                  label: const Text("اختيار ملف Excel"),
                ),
                const SizedBox(width: 12),
                if (_fileName != null)
                  Expanded(child: Text("الملف: $_fileName", overflow: TextOverflow.ellipsis)),
              ],
            ),
          ),

          if (_loading) const LinearProgressIndicator(),

          const SizedBox(height: 8),

          Expanded(
            child: !hasData
                ? const Center(child: Text("لا توجد بيانات معروضة"))
                : Scrollbar(
              thumbVisibility: true,
              child: SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: SingleChildScrollView(
                  child: DataTable(
                    columns: _excelData.first
                        .map((cell) => DataColumn(
                      label: Text(
                        cell,
                        style: const TextStyle(fontWeight: FontWeight.bold),
                      ),
                    ))
                        .toList(),
                    rows: _excelData.skip(1).map((row) {
                      return DataRow(
                        cells: row
                            .map((cell) => DataCell(Text(cell)))
                            .toList(),
                      );
                    }).toList(),
                  ),
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }
}
