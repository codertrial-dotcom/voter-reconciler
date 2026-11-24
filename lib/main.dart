import 'dart:convert';
import 'dart:io';
import 'dart:math';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:http/http.dart' as http;
import 'package:excel/excel.dart' as xls;
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';
import 'package:permission_handler/permission_handler.dart';

void main() {
  runApp(const VoterReconcilerApp());
}

class VoterReconcilerApp extends StatelessWidget {
  const VoterReconcilerApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Electoral Roll Reconciler',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: const Color(0xFF1A5276)),
        useMaterial3: true,
      ),
      home: const ReconcilerDashboard(),
    );
  }
}

// ==========================================
// DATA MODEL
// ==========================================
class VoterRecord {
  final String wardOrMajra;
  final String voterName;
  final String relativeName;
  final String relationType;
  final int age;
  final String gender;
  String status;

  VoterRecord({
    required this.wardOrMajra,
    required this.voterName,
    required this.relativeName,
    required this.relationType,
    required this.age,
    required this.gender,
    this.status = 'Unverified',
  });

  factory VoterRecord.fromJson(Map<String, dynamic> json) {
    return VoterRecord(
      wardOrMajra: (json['ward_majra'] ?? 'N/A').toString().trim(),
      voterName: (json['voter_name'] ?? '').toString().trim(),
      relativeName: (json['relative_name'] ?? '').toString().trim(),
      relationType: (json['relation_type'] ?? 'Father/Husband').toString().trim(),
      age: int.tryParse(json['age']?.toString() ?? '0') ?? 0,
      gender: _normalizeGender((json['gender'] ?? 'M').toString()),
    );
  }

  static String _normalizeGender(String raw) {
    final g = raw.trim().toLowerCase();
    if (g.startsWith('f') || g.contains('महिला') || g.contains('स्त्री')) {
      return 'F';
    }
    if (g.startsWith('m') || g.contains('पुरुष') || g.contains('पु')) {
      return 'M';
    }
    return 'O';
  }
}

// ==========================================
// FUZZY MATCHING ENGINE (LEVENSHTEIN + TOKEN SORT)
// ==========================================
class FuzzyMatcher {
  /// Normalizes Hindi (Devanagari) & English strings by removing honorifics & punctuation
  static String normalizeText(String input) {
    var cleaned = input.toLowerCase();
    final honorifics = [
      'shri', 'smt', 'kumari', 'late', 'mr', 'mrs',
      'श्री', 'श्रीमती', 'कुमारी', 'स्व.', 'स्व', 'मो.', 'मोहम्मद'
    ];
    for (var h in honorifics) {
      cleaned = cleaned.replaceAll(RegExp('\\b$h\\b'), '');
    }
    // Remove punctuation and normalize common Devanagari Bindi/Chandra-bindu variations
    cleaned = cleaned
        .replaceAll('ँ', 'ं')
        .replaceAll(RegExp(r'[^\w\s\u0900-\u097F]'), ' ')
        .replaceAll(RegExp(r'\s+'), ' ')
        .trim();
    return cleaned;
  }

  /// Calculates Token Sort Ratio (0.0 to 100.0)
  static double tokenSortRatio(String s1, String s2) {
    final norm1 = normalizeText(s1);
    final norm2 = normalizeText(s2);

    if (norm1.isEmpty && norm2.isEmpty) return 100.0;
    if (norm1.isEmpty || norm2.isEmpty) return 0.0;
    if (norm1 == norm2) return 100.0;

    final tokens1 = norm1.split(' ')..sort();
    final tokens2 = norm2.split(' ')..sort();

    final sorted1 = tokens1.join(' ');
    final sorted2 = tokens2.join(' ');

    final dist = _levenshteinDistance(sorted1, sorted2);
    final maxLen = max(sorted1.length, sorted2.length);
    return ((maxLen - dist) / maxLen) * 100.0;
  }

  static int _levenshteinDistance(String s, String t) {
    if (s == t) return 0;
    if (s.isEmpty) return t.length;
    if (t.isEmpty) return s.length;

    List<int> v0 = List<int>.generate(t.length + 1, (i) => i);
    List<int> v1 = List<int>.filled(t.length + 1, 0);

    for (int i = 0; i < s.length; i++) {
      v1[0] = i + 1;
      for (int j = 0; j < t.length; j++) {
        int cost = (s[i] == t[j]) ? 0 : 1;
        v1[j + 1] = min(v1[j] + 1, min(v0[j + 1] + 1, v0[j] + cost));
      }
      for (int j = 0; j <= t.length; j++) {
        v0[j] = v1[j];
      }
    }
    return v1[t.length];
  }

  /// Contextual Loose Match for Ward/Majra/Anubhaag
  static bool looseLocationMatch(String loc1, String loc2) {
    if (loc1 == 'N/A' || loc2 == 'N/A' || loc1.isEmpty || loc2.isEmpty) {
      return true; // Do not reject if PDF header lacked Majra info
    }
    final n1 = normalizeText(loc1);
    final n2 = normalizeText(loc2);
    if (n1.contains(n2) || n2.contains(n1)) return true;
    return tokenSortRatio(n1, n2) >= 50.0;
  }
}

// ==========================================
// MAIN DASHBOARD UI
// ==========================================
class ReconcilerDashboard extends StatefulWidget {
  const ReconcilerDashboard({super.key});

  @override
  State<ReconcilerDashboard> createState() => _ReconcilerDashboardState();
}

class _ReconcilerDashboardState extends State<ReconcilerDashboard> {
  final TextEditingController _apiKeyController = TextEditingController();

  List<PlatformFile> _panchayatFiles = [];
  List<PlatformFile> _vidhanSabhaFiles = [];

  // Configurable Rules
  double _ageTolerance = 2.0;
  double _similarityThreshold = 80.0;
  bool _enforceLooseMajraMatch = false;

  // State & Metrics
  bool _isProcessing = false;
  double _progress = 0.0;
  String _statusMessage = 'Upload PDFs and enter your Gemini API Key to start.';

  List<VoterRecord> _panchayatVoters = [];
  List<VoterRecord> _vidhanSabhaVoters = [];
  List<VoterRecord> _missingVoters = [];

  // Pick Multiple PDFs
  Future<void> _pickFiles(bool isPanchayat) async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['pdf'],
      allowMultiple: true,
    );

    if (result != null) {
      setState(() {
        if (isPanchayat) {
          _panchayatFiles = result.files;
        } else {
          _vidhanSabhaFiles = result.files;
        }
      });
    }
  }

  // Gemini 1.5 Flash PDF Extraction
  Future<List<VoterRecord>> _extractVotersFromPdf(
      PlatformFile pdfFile, String apiKey) async {
    if (pdfFile.path == null) return [];

    final bytes = await File(pdfFile.path!).readAsBytes();
    final base64Pdf = base64Encode(bytes);

    final url = Uri.parse(
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=$apiKey');

    const prompt = '''
You are an expert Indian Electoral Roll OCR parser. Extract all voter entries from this PDF (Hindi Devanagari or English).
Ignore deleted voters (marked with DELETED / विलोपित).
Return ONLY a valid JSON array of objects with these exact keys:
- "ward_majra": String (Ward number, Majra, or Anubhaag name/number from page header)
- "voter_name": String (Full name of voter in original script, or transliterated consistently)
- "relative_name": String (Father/Husband/Mother name)
- "relation_type": String ("Father", "Husband", "Mother", or "Other")
- "age": Integer (Voter age)
- "gender": String ("M", "F", or "O")
''';

    final payload = {
      "contents": [
        {
          "parts": [
            {"text": prompt},
            {
              "inline_data": {
                "mime_type": "application/pdf",
                "data": base64Pdf,
              }
            }
          ]
        }
      ],
      "generationConfig": {
        "response_mime_type": "application/json",
        "temperature": 0.1,
      }
    };

    final response = await http.post(
      url,
      headers: {'Content-Type': 'application/json'},
      body: jsonEncode(payload),
    );

    if (response.statusCode != 200) {
      throw Exception(
          'Gemini API Error (${response.statusCode}): ${response.body}');
    }

    final decoded = jsonDecode(response.body);
    final textResponse =
        decoded['candidates']?[0]?['content']?['parts']?[0]?['text'] ?? '[]';

    final List<dynamic> rawList = jsonDecode(textResponse);
    return rawList.map((e) => VoterRecord.fromJson(e)).toList();
  }

  // Main Execution Pipeline
  Future<void> _startReconciliation() async {
    final apiKey = _apiKeyController.text.trim();
    if (apiKey.isEmpty) {
      _showSnackBar('Please enter a valid Gemini API Key.');
      return;
    }
    if (_panchayatFiles.isEmpty || _vidhanSabhaFiles.isEmpty) {
      _showSnackBar(
          'Please select both Panchayat and Vidhan Sabha PDF files.');
      return;
    }

    setState(() {
      _isProcessing = true;
      _progress = 0.05;
      _panchayatVoters.clear();
      _vidhanSabhaVoters.clear();
      _missingVoters.clear();
      _statusMessage = 'Initializing Gemini 1.5 Flash OCR pipeline...';
    });

    try {
      final totalFiles = _panchayatFiles.length + _vidhanSabhaFiles.length;
      int processedFiles = 0;

      // Step 1: Parse Panchayat PDFs
      for (final file in _panchayatFiles) {
        setState(() {
          _statusMessage = 'Parsing Panchayat PDF: ${file.name}...';
        });
        final extracted = await _extractVotersFromPdf(file, apiKey);
        _panchayatVoters.addAll(extracted);
        processedFiles++;
        setState(() {
          _progress = (processedFiles / totalFiles) * 0.75;
        });
      }

      // Step 2: Parse Vidhan Sabha PDFs
      for (final file in _vidhanSabhaFiles) {
        setState(() {
          _statusMessage = 'Parsing Vidhan Sabha PDF: ${file.name}...';
        });
        final extracted = await _extractVotersFromPdf(file, apiKey);
        _vidhanSabhaVoters.addAll(extracted);
        processedFiles++;
        setState(() {
          _progress = (processedFiles / totalFiles) * 0.75;
        });
      }

      // Step 3: Local Fuzzy Matching Cross-Verification
      setState(() {
        _statusMessage =
            'Running local Levenshtein & Token Sort fuzzy matching...';
      });

      final List<VoterRecord> missing = [];
      final int totalPanchayat = _panchayatVoters.length;

      for (int i = 0; i < totalPanchayat; i++) {
        final pVoter = _panchayatVoters[i];
        bool matchFound = false;

        for (final vsVoter in _vidhanSabhaVoters) {
          // 1. Exact Gender Match
          if (pVoter.gender != vsVoter.gender) continue;

          // 2. Configurable Age Range Match (± years)
          if (pVoter.age > 0 && vsVoter.age > 0) {
            if ((pVoter.age - vsVoter.age).abs() > _ageTolerance.round()) {
              continue;
            }
          }

          // 3. Optional Contextual Majra/Ward Loose Match
          if (_enforceLooseMajraMatch) {
            if (!FuzzyMatcher.looseLocationMatch(
                pVoter.wardOrMajra, vsVoter.wardOrMajra)) {
              continue;
            }
          }

          // 4. Fuzzy Name & Relative Name Match
          final nameScore = FuzzyMatcher.tokenSortRatio(
              pVoter.voterName, vsVoter.voterName);
          final relScore = FuzzyMatcher.tokenSortRatio(
              pVoter.relativeName, vsVoter.relativeName);

          // Combined weighted score (60% Voter Name, 40% Relative Name)
          final combinedScore = (nameScore * 0.6) + (relScore * 0.4);

          if (nameScore >= _similarityThreshold &&
              combinedScore >= _similarityThreshold) {
            matchFound = true;
            break;
          }
        }

        if (!matchFound) {
          pVoter.status = 'Missing in Vidhan Sabha Roll';
          missing.add(pVoter);
        }

        if (i % 10 == 0 && totalPanchayat > 0) {
          setState(() {
            _progress = 0.75 + ((i / totalPanchayat) * 0.25);
          });
          await Future.delayed(const Duration(milliseconds: 1));
        }
      }

      setState(() {
        _missingVoters = missing;
        _progress = 1.0;
        _isProcessing = false;
        _statusMessage =
            'Reconciliation complete! Found ${_missingVoters.length} missing voters.';
      });
    } catch (e) {
      setState(() {
        _isProcessing = false;
        _statusMessage = 'Error: $e';
      });
      _showSnackBar('Error during processing: $e');
    }
  }

  // Export Missing Voters to Excel (.xlsx)
  Future<void> _exportToExcel() async {
    if (_missingVoters.isEmpty) {
      _showSnackBar('No missing voters to export!');
      return;
    }

    try {
      var excel = xls.Excel.createExcel();
      final String defaultSheet = excel.getDefaultSheet() ?? 'Sheet1';
      excel.rename(defaultSheet, 'Missing Voters');
      xls.Sheet sheetObject = excel['Missing Voters'];

      // Header Row
      final headers = [
        'Ward/Majra No.',
        'Voter Name',
        'Relative Name',
        'Relation Type',
        'Age',
        'Gender',
        'Status'
      ];
      sheetObject.appendRow(headers.map((h) => xls.TextCellValue(h)).toList());

      // Data Rows
      for (final v in _missingVoters) {
        sheetObject.appendRow([
          xls.TextCellValue(v.wardOrMajra),
          xls.TextCellValue(v.voterName),
          xls.TextCellValue(v.relativeName),
          xls.TextCellValue(v.relationType),
          xls.IntCellValue(v.age),
          xls.TextCellValue(v.gender),
          xls.TextCellValue(v.status),
        ]);
      }

      final fileBytes = excel.save();
      if (fileBytes == null) return;

      // Attempt direct save to Android Downloads folder
      Directory? directory;
      if (Platform.isAndroid) {
        await Permission.storage.request();
        directory = Directory('/storage/emulated/0/Download');
        if (!await directory.exists()) {
          directory = await getExternalStorageDirectory();
        }
      }
      directory ??= await getApplicationDocumentsDirectory();

      final String filePath = '${directory.path}/missing_panchayat_voters.xlsx';
      final File outFile = File(filePath);
      await outFile.writeAsBytes(fileBytes, flush: true);

      _showSnackBar('Saved to $filePath');

      // Also trigger native share/save dialog for Android 11+ Scoped Storage convenience
      await Share.shareXFiles(
        [XFile(filePath)],
        text: 'Missing Panchayat Voters Reconciliation Report',
      );
    } catch (e) {
      _showSnackBar('Failed to export Excel: $e');
    }
  }

  void _showSnackBar(String msg) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(msg), behavior: SnackBarBehavior.floating),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Electoral Roll Reconciler'),
        backgroundColor: Theme.of(context).colorScheme.primary,
        foregroundColor: Colors.white,
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            // API Key Input
            Card(
              elevation: 2,
              child: Padding(
                padding: const EdgeInsets.all(12),
                child: TextField(
                  controller: _apiKeyController,
                  obscureText: true,
                  decoration: const InputDecoration(
                    labelText: 'Gemini API Key (Gemini 1.5 Flash)',
                    hintText: 'AIzaSy...',
                    prefixIcon: Icon(Icons.key),
                    border: OutlineInputBorder(),
                  ),
                ),
              ),
            ),
            const SizedBox(height: 12),

            // 1. Dual Multi-PDF Picker
            Row(
              children: [
                Expanded(
                  child: _buildFilePickerCard(
                    title: 'Panchayat List(s)',
                    count: _panchayatFiles.length,
                    icon: Icons.account_balance,
                    onTap: _isProcessing ? null : () => _pickFiles(true),
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: _buildFilePickerCard(
                    title: 'Vidhan Sabha List(s)',
                    count: _vidhanSabhaFiles.length,
                    icon: Icons.how_to_vote,
                    onTap: _isProcessing ? null : () => _pickFiles(false),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 12),

            // 2. Configuration & Matching Rules
            Card(
              elevation: 2,
              child: Padding(
                padding: const EdgeInsets.all(14),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'Fuzzy Matching Rules',
                      style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold),
                    ),
                    const SizedBox(height: 8),
                    Text('Age Tolerance: ±${_ageTolerance.round()} years'),
                    Slider(
                      value: _ageTolerance,
                      min: 0,
                      max: 10,
                      divisions: 10,
                      label: '±${_ageTolerance.round()} yrs',
                      onChanged: _isProcessing
                          ? null
                          : (val) => setState(() => _ageTolerance = val),
                    ),
                    Text(
                        'Name & Relative Similarity Threshold: ${_similarityThreshold.round()}%'),
                    Slider(
                      value: _similarityThreshold,
                      min: 50,
                      max: 100,
                      divisions: 50,
                      label: '${_similarityThreshold.round()}%',
                      onChanged: _isProcessing
                          ? null
                          : (val) => setState(() => _similarityThreshold = val),
                    ),
                    SwitchListTile(
                      contentPadding: EdgeInsets.zero,
                      title: const Text('Enforce Loose Ward/Majra Match'),
                      subtitle: const Text('Gender is strictly matched (Exact)'),
                      value: _enforceLooseMajraMatch,
                      onChanged: _isProcessing
                          ? null
                          : (val) =>
                              setState(() => _enforceLooseMajraMatch = val),
                    ),
                  ],
                ),
              ),
            ),
            const SizedBox(height: 12),

            // 3. Execution Button & Real-Time Progress Bar
            ElevatedButton.icon(
              onPressed: _isProcessing ? null : _startReconciliation,
              icon: const Icon(Icons.play_arrow),
              label: const Padding(
                padding: EdgeInsets.symmetric(vertical: 14),
                child: Text(
                  'Start Reconciliation',
                  style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold),
                ),
              ),
              style: ElevatedButton.styleFrom(
                backgroundColor: Theme.of(context).colorScheme.primary,
                foregroundColor: Colors.white,
              ),
            ),
            const SizedBox(height: 12),

            if (_isProcessing || _progress > 0) ...[
              LinearProgressIndicator(value: _progress, minHeight: 8),
              const SizedBox(height: 6),
              Text(
                _statusMessage,
                textAlign: TextAlign.center,
                style: const TextStyle(fontSize: 13, color: Colors.black87),
              ),
              const SizedBox(height: 12),
            ],

            // Summary Cards
            Row(
              children: [
                Expanded(
                  child: _buildSummaryCard(
                    'Panchayat Parsed',
                    _panchayatVoters.length.toString(),
                    Colors.blue.shade700,
                  ),
                ),
                const SizedBox(width: 8),
                Expanded(
                  child: _buildSummaryCard(
                    'Vidhan Sabha Parsed',
                    _vidhanSabhaVoters.length.toString(),
                    Colors.teal.shade700,
                  ),
                ),
                const SizedBox(width: 8),
                Expanded(
                  child: _buildSummaryCard(
                    'Missing Flagged',
                    _missingVoters.length.toString(),
                    Colors.red.shade700,
                  ),
                ),
              ],
            ),
            const SizedBox(height: 16),

            // 4. One-Tap Excel Export Button
            FilledButton.icon(
              onPressed:
                  (_missingVoters.isNotEmpty && !_isProcessing) ? _exportToExcel : null,
              icon: const Icon(Icons.table_view),
              label: const Padding(
                padding: EdgeInsets.symmetric(vertical: 14),
                child: Text(
                  'Export Missing Voters to Excel (.xlsx)',
                  style: TextStyle(fontSize: 15, fontWeight: FontWeight.bold),
                ),
              ),
              style: FilledButton.styleFrom(
                backgroundColor: Colors.green.shade700,
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildFilePickerCard({
    required String title,
    required int count,
    required IconData icon,
    required VoidCallback? onTap,
  }) {
    return InkWell(
      onTap: onTap,
      child: Card(
        elevation: 2,
        child: Padding(
          padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 12),
          child: Column(
            children: [
              Icon(icon, size: 32, color: Theme.of(context).colorScheme.primary),
              const SizedBox(height: 8),
              Text(
                title,
                textAlign: TextAlign.center,
                style: const TextStyle(fontWeight: FontWeight.bold),
              ),
              const SizedBox(height: 4),
              Text(
                '$count PDF(s) selected',
                style: TextStyle(color: Colors.grey.shade700, fontSize: 12),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildSummaryCard(String label, String value, Color color) {
    return Card(
      elevation: 2,
      child: Padding(
        padding: const EdgeInsets.symmetric(vertical: 14, horizontal: 8),
        child: Column(
          children: [
            Text(
              value,
              style: TextStyle(
                fontSize: 22,
                fontWeight: FontWeight.bold,
                color: color,
              ),
            ),
            const SizedBox(height: 4),
            Text(
              label,
              textAlign: TextAlign.center,
              style: const TextStyle(fontSize: 12),
            ),
          ],
        ),
      ),
    );
  }
}
