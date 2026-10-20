const SHEET_STUDENTS = 'Siswa';
const SHEET_EXAMS = 'Daftar Ujian';
const STUDENT_HEADERS = ['Kelas', 'No Absen', 'Nama'];
const EXAM_HEADERS = ['Kode', 'Nama Ujian', 'Token', 'Kelas', 'Buka', 'Tutup', 'Durasi (menit)', 'Tampilkan Nilai', 'Tampilkan Pembahasan', 'Maks Pelanggaran', 'Jeda (detik)', 'Acak Soal', 'Acak Opsi', 'Status'];
const QUESTION_HEADERS = ['No', 'Tipe', 'Pertanyaan', 'Gambar', 'A', 'B', 'C', 'D', 'E', 'Kunci', 'Bobot', 'Pembahasan'];
const PARTICIPANT_HEADERS = ['Kelas', 'No Absen', 'Nama', 'Status', 'Login Terakhir', 'Mulai', 'Selesai', 'Pelanggaran', 'Riwayat Pelanggaran', 'Skor Otomatis', 'Skor Manual', 'Total Skor', 'Skor Maks', 'Nilai', 'Sesi', 'Jawaban Sementara', 'Data'];
const ANSWER_HEADERS = ['Waktu', 'Kelas', 'No Absen', 'Nama', 'No', 'Tipe', 'Pertanyaan', 'Jawaban', 'Kunci', 'Benar', 'Skor', 'Bobot'];
const WIDE_ANSWER_BASE_HEADERS = ['Waktu Selesai', 'Kelas', 'No Absen', 'Nama'];
const CACHE_SECONDS = 600;

// Aksi guru memakai akhiran garis bawah agar tidak bisa dipanggil dari halaman siswa.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📝 Ujian')
    .addItem('Siapkan spreadsheet', 'setupSpreadsheet_')
    .addItem('Buat ujian baru', 'createExam_')
    .addSeparator()
    .addItem('Reset siswa', 'resetStudents_')
    .addItem('Arsipkan ujian', 'archiveExam_')
    .addItem('Bersihkan cache', 'clearAppCache_')
    .addToUi();
}

function migrateParticipantHeaders_(sheet) {
  if (sheet.getLastColumn() >= 11 && String(sheet.getRange(1, 11).getDisplayValue()).trim() === 'Skor Essay') {
    sheet.getRange(1, 11).setValue('Skor Manual');
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SiMatik')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// API JSON untuk halaman PHP lokal; spreadsheet tetap diakses lewat Apps Script.
function doPost(e) {
  try {
    const payload = JSON.parse(e && e.postData ? e.postData.contents : '{}');
    const args = Array.isArray(payload.args) ? payload.args : [];
    let result;
    switch (payload.action) {
      case 'getLoginOptions':
        result = getLoginOptions();
        break;
      case 'loginStudent':
        result = loginStudent(args[0] || {});
        break;
      case 'startExam':
        result = startExam(args[0], args[1]);
        break;
      case 'saveAnswers':
        result = saveAnswers(args[0], args[1], args[2]);
        break;
      case 'reportViolation':
        result = reportViolation(args[0], args[1], args[2]);
        break;
      case 'submitExam':
        result = submitExam(args[0], args[1], args[2]);
        break;
      default:
        throw new Error('Aksi tidak dikenal.');
    }
    return jsonOutput_({ ok: true, result });
  } catch (error) {
    console.error(error);
    return jsonOutput_({ ok: false, error: String(error && error.message || 'Terjadi kendala di server.') });
  }
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLoginOptions() {
  const names = cachedTable_('web:students', SHEET_STUDENTS)
    .map(row => String(row.Kelas || '').trim()).filter(Boolean);
  return [...new Set(names)].sort();
}

// Membuat struktur awal tanpa menimpa data yang sudah ada.
function setupSpreadsheet_() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Buka spreadsheet terlebih dahulu, lalu jalankan menu ini.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ensureSheet_(ss, SHEET_STUDENTS, STUDENT_HEADERS);
  ensureSheet_(ss, SHEET_EXAMS, EXAM_HEADERS);
  const students = ss.getSheetByName(SHEET_STUDENTS);
  const exams = ss.getSheetByName(SHEET_EXAMS);
  if (students.getLastRow() < 2) {
    students.getRange(2, 1, 4, 3).setValues([
      ['8A', 1, 'Andi Pratama'], ['8A', 2, 'Bunga Lestari'],
      ['8B', 1, 'Citra Maharani'], ['8B', 2, 'Dimas Saputra']
    ]);
  }
  const hasSampleExam = table_(exams).rows.some(row => String(row.Kode || '').trim().toUpperCase() === 'KUIS1');
  if (!hasSampleExam) {
    exams.appendRow(['KUIS1', 'Kuis Matematika 1', 'MTK8', '8A,8B', '', '', 20,
      'Langsung', 'Ya', 3, 60, 'Ya', 'Ya', 'Aktif']);
  }
  applyBaseFormatting_(students, exams);
  ensureQuestionSheet_(ss, 'KUIS1', true);
  ensureParticipantSheets_(ss, 'KUIS1');
  ensureAnswerEditTrigger_(ss);
  clearCache_();
  SpreadsheetApp.getUi().alert('Spreadsheet SiMatik siap digunakan. Contoh ujian KUIS1 sudah tersedia.');
}

function ensureAnswerEditTrigger_(spreadsheet) {
  const spreadsheetId = spreadsheet.getId();
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(trigger =>
    trigger.getHandlerFunction() === 'handleSpreadsheetEdit_' &&
    trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
    trigger.getTriggerSourceId() === spreadsheetId);
  if (!hasTrigger) {
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEdit' &&
          trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
          trigger.getTriggerSourceId() === spreadsheetId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    ScriptApp.newTrigger('handleSpreadsheetEdit_').forSpreadsheet(spreadsheet).onEdit().create();
  }
}

function createExam_() {
  const ui = SpreadsheetApp.getUi();
  const codePrompt = ui.prompt('Buat ujian baru', 'Masukkan kode ujian (huruf/angka tanpa spasi):', ui.ButtonSet.OK_CANCEL);
  if (codePrompt.getSelectedButton() !== ui.Button.OK) return;
  const code = codePrompt.getResponseText().trim().toUpperCase();
  if (!/^[A-Z0-9]{1,92}$/.test(code)) throw new Error('Kode ujian hanya boleh berisi 1–92 huruf atau angka.');
  const namePrompt = ui.prompt('Buat ujian baru', 'Masukkan nama ujian:', ui.ButtonSet.OK_CANCEL);
  if (namePrompt.getSelectedButton() !== ui.Button.OK) return;
  const name = namePrompt.getResponseText().trim();
  if (!name) throw new Error('Nama ujian tidak boleh kosong.');
  const ss = getSpreadsheet_();
  const examSheet = ss.getSheetByName(SHEET_EXAMS);
  if (!examSheet) throw new Error('Jalankan "Siapkan spreadsheet" terlebih dahulu.');
  if (examByCode_(code)) throw new Error('Kode ujian tersebut sudah digunakan.');
  ensureQuestionSheet_(ss, code, false);
  ensureParticipantSheets_(ss, code);
  const token = uniqueToken_();
  examSheet.appendRow([code, name, token, 'SEMUA', '', '', 60, 'Tidak', 'Tidak', 0, 60, 'Tidak', 'Tidak', 'Nonaktif']);
  applyExamValidation_(examSheet);
  clearCache_();
  ui.alert('Ujian dibuat. Kode: ' + code + '\nToken siswa: ' + token + '\nStatus awal: Nonaktif.');
}

function resetStudents_() {
  const ui = SpreadsheetApp.getUi();
  const code = ui.prompt('Reset siswa', 'Kode ujian:', ui.ButtonSet.OK_CANCEL);
  if (code.getSelectedButton() !== ui.Button.OK) return;
  const classroom = ui.prompt('Reset siswa', 'Kelas siswa:', ui.ButtonSet.OK_CANCEL);
  if (classroom.getSelectedButton() !== ui.Button.OK) return;
  const numbers = ui.prompt('Reset siswa', 'Nomor absen (pisahkan dengan koma):', ui.ButtonSet.OK_CANCEL);
  if (numbers.getSelectedButton() !== ui.Button.OK) return;
  const examCode = code.getResponseText().trim().toUpperCase();
  const className = classroom.getResponseText().trim();
  const wanted = numbers.getResponseText().split(',').map(normalizeNumber_).filter(Boolean);
  if (!wanted.length) throw new Error('Masukkan setidaknya satu nomor absen.');
  const ss = getSpreadsheet_();
  const participantSheet = ss.getSheetByName('Peserta_' + examCode);
  const answerSheet = ss.getSheetByName('Jawaban_' + examCode);
  const detailSheet = ss.getSheetByName(detailAnswerSheetName_(examCode));
  if (!participantSheet || !answerSheet || !detailSheet) throw new Error('Sheet ujian tidak ditemukan.');
  const p = table_(participantSheet);
  const rowsToDelete = [];
  p.rows.forEach(row => {
    if (String(row.Kelas).trim() === className && wanted.includes(normalizeNumber_(row['No Absen']))) rowsToDelete.push(row.__rowNumber);
  });
  rowsToDelete.sort((a, b) => b - a).forEach(row => participantSheet.deleteRow(row));
  const answerTable = table_(answerSheet);
  const wideRows = [];
  answerTable.rows.forEach(row => {
    if (String(row.Kelas).trim() === className && wanted.includes(normalizeNumber_(row['No Absen']))) wideRows.push(row.__rowNumber);
  });
  wideRows.sort((x, y) => y - x).forEach(row => answerSheet.deleteRow(row));
  const a = table_(detailSheet);
  const detailRows = [];
  a.rows.forEach(row => {
    if (String(row.Kelas).trim() === className && wanted.includes(normalizeNumber_(row['No Absen']))) detailRows.push(row.__rowNumber);
  });
  detailRows.sort((x, y) => y - x).forEach(row => detailSheet.deleteRow(row));
  ui.alert('Reset selesai. Baris peserta yang dihapus: ' + rowsToDelete.length + '.');
}

function archiveExam_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Arsipkan ujian', 'Masukkan kode ujian:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const code = response.getResponseText().trim().toUpperCase();
  const ss = getSpreadsheet_();
  const exam = examByCode_(code);
  if (!exam) throw new Error('Ujian tidak ditemukan.');
  const names = ['Soal_' + code, 'Peserta_' + code, 'Jawaban_' + code, detailAnswerSheetName_(code)];
  const sourceSheets = names.map(name => ss.getSheetByName(name));
  if (sourceSheets.some(sheet => !sheet)) throw new Error('Salah satu sheet ujian tidak ditemukan.');
  const properties = PropertiesService.getScriptProperties();
  let archiveId = properties.getProperty('ARSIP_UJIAN_ID');
  let archive;
  if (archiveId) {
    archive = SpreadsheetApp.openById(archiveId);
  } else {
    archive = SpreadsheetApp.create('Arsip Ujian');
    archiveId = archive.getId();
    properties.setProperty('ARSIP_UJIAN_ID', archiveId);
    const first = archive.getSheets()[0];
    first.setName('Arsip');
    first.getRange('A1').setValue('Sheet arsip ujian akan muncul di sini.');
  }
  sourceSheets.forEach(sheet => {
    const copied = sheet.copyTo(archive).setName(uniqueSheetName_(archive, sheet.getName()));
    const values = sheet.getDataRange().getValues();
    copied.getDataRange().setValues(values);
  });
  const placeholder = archive.getSheetByName('Arsip');
  if (placeholder && archive.getSheets().length > 1 && placeholder.getRange('A1').getValue() === 'Sheet arsip ujian akan muncul di sini.') archive.deleteSheet(placeholder);
  sourceSheets.forEach(sheet => ss.deleteSheet(sheet));
  writeExamField_(exam, 'Status', 'Diarsipkan');
  clearCache_();
  ui.alert('Ujian ' + code + ' berhasil diarsipkan.\nID file arsip: ' + archiveId);
}

function clearAppCache_() {
  clearCache_();
  SpreadsheetApp.getUi().alert('Cache aplikasi berhasil dibersihkan.');
}

function clearCache_() {
  CacheService.getScriptCache().removeAll(['web:students', 'web:exams']);
  getSpreadsheet_().getSheets().forEach(sheet => {
    if (/^Soal_/.test(sheet.getName())) CacheService.getScriptCache().remove('web:questions:' + sheet.getName().slice(5));
  });
}

// Cache dipangkas berdasarkan sheet yang diedit guru.
function handleSpreadsheetEdit_(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  const cache = CacheService.getScriptCache();
  if (name === SHEET_STUDENTS) cache.remove('web:students');
  else if (name === SHEET_EXAMS) cache.remove('web:exams');
  else if (/^Soal_/.test(name)) {
    const code = name.slice(5);
    cache.remove('web:questions:' + code);
    const editsQuestionLayout = e.range.getColumn() <= 2 && e.range.getLastColumn() >= 1;
    const editsWeight = e.range.getColumn() <= 11 && e.range.getLastColumn() >= 11;
    if (e.range.getRow() > 1 && (editsQuestionLayout || editsWeight)) {
      const ss = getSpreadsheet_();
      const questions = table_(e.range.getSheet()).rows.filter(row => row.No !== '' && row.No != null);
      ensureAnswerSheets_(ss, code, questions);
      const participants = ss.getSheetByName('Peserta_' + code);
      if (participants) {
        table_(participants).rows.forEach(row => {
          if (String(row.Kelas || '').trim()) applyParticipantFormulas_(participants, row.__rowNumber, code);
        });
      }
    }
  }
  else if (/^Jawaban_/.test(name)) syncManualScores_(e);
}

// Memeriksa identitas, jadwal, kelas, dan sesi perangkat siswa.
function loginStudent(input) {
  const classroom = String(input.classroom || '').trim();
  const number = normalizeNumber_(input.number);
  const token = String(input.token || '').trim();
  if (!classroom || !number || !token) throw new Error('Lengkapi kelas, nomor absen, dan token.');
  const student = cachedTable_('web:students', SHEET_STUDENTS).find(row =>
    String(row.Kelas).trim() === classroom && normalizeNumber_(row['No Absen']) === number);
  if (!student) throw new Error('Data siswa tidak ditemukan. Periksa kelas dan nomor absen.');
  const exam = cachedTable_('web:exams', SHEET_EXAMS).find(row => String(row.Token || '').trim() === token);
  if (!exam) throw new Error('Token ujian tidak ditemukan.');
  const now = Date.now();
  if (String(exam.Status) !== 'Aktif') throw new Error('Ujian ini belum aktif.');
  if (!classAllowed_(exam.Kelas, classroom)) throw new Error('Kelasmu tidak terdaftar untuk ujian ini.');
  const code = String(exam.Kode).trim().toUpperCase();
  const questions = getQuestions_(code);
  const ss = getSpreadsheet_();
  const participantSheet = ss.getSheetByName('Peserta_' + code);
  if (!participantSheet) throw new Error('Data peserta ujian belum disiapkan. Jalankan "Siapkan spreadsheet".');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureAnswerSheets_(ss, code, questions);
    migrateParticipantHeaders_(participantSheet);
    const t = table_(participantSheet);
    const existing = t.rows.find(row => String(row.Kelas).trim() === classroom && normalizeNumber_(row['No Absen']) === number);
    let rowNumber = existing ? existing.__rowNumber : 0;
    let participant;
    if (!existing) {
      if (asMillis_(exam.Buka) && now < asMillis_(exam.Buka)) throw new Error('Ujian belum dibuka.');
      if (asMillis_(exam.Tutup) && now > asMillis_(exam.Tutup)) throw new Error('Waktu ujian sudah ditutup.');
      rowNumber = firstEmptyParticipantRow_(participantSheet);
      participant = { Kelas: classroom, 'No Absen': Number(number), Nama: String(student.Nama), Status: 'Login',
        'Login Terakhir': new Date(), Pelanggaran: 0, 'Riwayat Pelanggaran': '', 'Jawaban Sementara': '{}', Data: '{}' };
      participantSheet.getRange(rowNumber, 1, 1, PARTICIPANT_HEADERS.length).setValues([[
        classroom, Number(number), String(student.Nama), 'Login', new Date(), '', '', 0, '', '', '', '', '', '', '', '{}', '{}'
      ]]);
      applyParticipantFormulas_(participantSheet, rowNumber, code);
    } else {
      participant = existing;
      if (String(participant.Status) === 'Selesai') {
        return { state: 'alreadyCompleted', name: String(exam['Nama Ujian']), studentName: String(student.Nama) };
      }
      if (asMillis_(exam.Buka) && now < asMillis_(exam.Buka)) throw new Error('Ujian belum dibuka.');
      if (asMillis_(exam.Tutup) && now > asMillis_(exam.Tutup)) throw new Error('Waktu ujian sudah ditutup.');
    }
    const status = String(participant.Status || 'Login');
    const session = Utilities.getUuid();
    const data = parseJson_(participant.Data, {});
    participantSheet.getRange(rowNumber, column_(t.headers, 'Login Terakhir')).setValue(new Date());
    participantSheet.getRange(rowNumber, column_(t.headers, 'Sesi')).setValue(session);
    const common = { code, name: String(exam['Nama Ujian']), studentName: String(student.Nama), classroom, number: Number(number),
      durationMinutes: durationMinutes_(exam), maxViolations: maxViolations_(exam),
      pauseSeconds: pauseSeconds_(exam), questionCount: questions.length, session };
    if (status === 'Mengerjakan') {
      if (!data.deadline) throw new Error('Data sesi ujian tidak lengkap. Hubungi guru.');
      return Object.assign(common, { state: 'exam', questions: publicQuestions_(questions, data), answers: parseJson_(participant['Jawaban Sementara'], {}),
        deadline: Number(data.deadline), pauseUntil: Number(data.pauseUntil) || 0, violations: Number(participant.Pelanggaran) || 0,
        violationHistory: String(participant['Riwayat Pelanggaran'] || ''), serverNow: Date.now() });
    }
    return Object.assign(common, { state: 'rules' });
  } finally {
    lock.releaseLock();
  }
}

// Membuat urutan soal/opsi dan deadline sekali untuk setiap sesi ujian.
function startExam(code, session) {
  return withParticipant_(code, session, (ctx) => {
    if (String(ctx.participant.Status) === 'Selesai') {
      return { state: 'alreadyCompleted', name: String(ctx.exam['Nama Ujian']), studentName: String(ctx.participant.Nama) };
    }
    if (String(ctx.participant.Status) === 'Mengerjakan') return examState_(ctx);
    const now = Date.now();
    const closeAt = asMillis_(ctx.exam.Tutup);
    const duration = durationMinutes_(ctx.exam) * 60000;
    const deadline = closeAt ? Math.min(now + duration, closeAt) : now + duration;
    const questionOrder = ctx.questions.map((_, i) => i);
    if (isYes_(ctx.exam['Acak Soal'])) shuffle_(questionOrder);
    const optionOrder = {};
    ctx.questions.forEach((question, index) => {
      const available = ['A', 'B', 'C', 'D', 'E'].filter(key => String(question[key] || '').trim());
      if (isYes_(ctx.exam['Acak Opsi'])) shuffle_(available);
      optionOrder[String(index)] = available;
    });
    const data = { questionOrder, optionOrder, deadline, pauseUntil: 0, lastViolationAt: 0 };
    const sheet = ctx.sheet;
    sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Status')).setValue('Mengerjakan');
    sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Mulai')).setValue(new Date(now));
    sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Data')).setValue(JSON.stringify(data));
    sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Jawaban Sementara')).setValue('{}');
    SpreadsheetApp.flush();
    ctx.participant.Status = 'Mengerjakan';
    ctx.participant.Data = JSON.stringify(data);
    ctx.participant['Jawaban Sementara'] = '{}';
    return examState_(ctx);
  });
}

// Menyimpan jawaban sementara terpisah dari data pelanggaran.
function saveAnswers(code, session, answers) {
  return withParticipant_(code, session, ctx => {
    ensureExamActive_(ctx);
    const data = parseJson_(ctx.participant.Data, {});
    if (Date.now() > Number(data.deadline) + 60000) {
      finishParticipant_(ctx, parseJson_(ctx.participant['Jawaban Sementara'], {}));
      return { state: 'finished', result: resultFor_(ctx.exam, code, ctx.participant) };
    }
    const normalized = {};
    ctx.questions.forEach(q => {
      const val = answers && Object.prototype.hasOwnProperty.call(answers, String(q.No)) ? answers[String(q.No)] : '';
      normalized[String(q.No)] = String(val == null ? '' : val).slice(0, 5000);
    });
    ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Jawaban Sementara')).setValue(JSON.stringify(normalized));
    return { saved: true, serverNow: Date.now() };
  });
}

// Mencatat pelanggaran dengan jeda dan deduplikasi sisi server.
function reportViolation(code, session, reason) {
  return withParticipant_(code, session, ctx => {
    ensureExamActive_(ctx);
    const now = Date.now();
    const data = parseJson_(ctx.participant.Data, {});
    if (now - (Number(data.lastViolationAt) || 0) < 3000) {
      return { duplicate: true, violations: Number(ctx.participant.Pelanggaran) || 0, pauseUntil: Number(data.pauseUntil) || 0, finished: false };
    }
    const violations = (Number(ctx.participant.Pelanggaran) || 0) + 1;
    const pauseUntil = now + pauseSeconds_(ctx.exam) * 1000;
    data.lastViolationAt = now;
    data.pauseUntil = pauseUntil;
    const history = String(ctx.participant['Riwayat Pelanggaran'] || '');
    const event = Utilities.formatDate(new Date(now), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') + ' - ' + String(reason || 'Pelanggaran').slice(0, 100);
    ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Pelanggaran')).setValue(violations);
    ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Riwayat Pelanggaran')).setValue(history ? history + '\n' + event : event);
    ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Data')).setValue(JSON.stringify(data));
    const limit = maxViolations_(ctx.exam);
    if (limit > 0 && violations >= limit) {
      finishParticipant_(ctx, parseJson_(ctx.participant['Jawaban Sementara'], {}));
      return { violations, pauseUntil, finished: true, result: resultFor_(ctx.exam, code, ctx.participant) };
    }
    return { violations, pauseUntil, finished: false };
  });
}

function submitExam(code, session, answers) {
  return withParticipant_(code, session, ctx => {
    if (String(ctx.participant.Status) === 'Selesai') return { result: resultFor_(ctx.exam, code, ctx.participant) };
    const data = parseJson_(ctx.participant.Data, {});
    if (Date.now() > Number(data.deadline) + 60000) answers = parseJson_(ctx.participant['Jawaban Sementara'], {});
    finishParticipant_(ctx, answers || parseJson_(ctx.participant['Jawaban Sementara'], {}));
    return { result: resultFor_(ctx.exam, code, ctx.participant) };
  });
}

function getExamResult(code, session) {
  return withParticipant_(code, session, ctx => {
    if (String(ctx.participant.Status) !== 'Selesai') throw new Error('Ujian belum dikumpulkan.');
    return resultFor_(ctx.exam, code, ctx.participant);
  });
}

function withParticipant_(code, session, callback) {
  const examCode = String(code || '').trim().toUpperCase();
  const exam = examByCode_(examCode);
  if (!exam) throw new Error('Ujian tidak ditemukan.');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName('Peserta_' + examCode);
  if (!sheet) throw new Error('Data peserta ujian tidak ditemukan.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(sheet);
    const participant = t.rows.find(row => String(row.Sesi || '') === String(session || ''));
    if (!participant) throw appError_('SESI_LAIN', 'SESI_LAIN: Akunmu login di perangkat lain. Silakan masuk kembali.');
    const rowNumber = participant.__rowNumber;
    if (String(participant.Sesi) !== String(session)) throw appError_('SESI_LAIN', 'SESI_LAIN: Akunmu login di perangkat lain. Silakan masuk kembali.');
    const ctx = { exam, code: examCode, sheet, headers: t.headers, rowNumber, participant, questions: getQuestions_(examCode) };
    return callback(ctx);
  } finally {
    lock.releaseLock();
  }
}

function examState_(ctx) {
  const data = parseJson_(ctx.participant.Data, {});
  return { state: 'exam', questions: publicQuestions_(ctx.questions, data),
    answers: parseJson_(ctx.participant['Jawaban Sementara'], {}), deadline: Number(data.deadline),
    pauseUntil: Number(data.pauseUntil) || 0, violations: Number(ctx.participant.Pelanggaran) || 0,
    violationHistory: String(ctx.participant['Riwayat Pelanggaran'] || ''), serverNow: Date.now() };
}

function ensureExamActive_(ctx) {
  if (String(ctx.participant.Status) !== 'Mengerjakan') throw new Error('Ujian sudah tidak dapat dikerjakan.');
}

// Menilai PG otomatis, lalu menyiapkan skor ISIAN/ESSAY untuk guru.
function finishParticipant_(ctx, answers) {
  const ss = getSpreadsheet_();
  ensureAnswerSheets_(ss, ctx.code, ctx.questions);
  const answerSheet = ss.getSheetByName('Jawaban_' + ctx.code);
  const detailSheet = ss.getSheetByName(detailAnswerSheetName_(ctx.code));
  if (!answerSheet || !detailSheet) throw new Error('Sheet jawaban tidak ditemukan.');
  const className = String(ctx.participant.Kelas);
  const number = normalizeNumber_(ctx.participant['No Absen']);
  const t = table_(detailSheet);
  const oldRows = [];
  t.rows.forEach(row => {
    if (String(row.Kelas) === className && normalizeNumber_(row['No Absen']) === number) oldRows.push(row.__rowNumber);
  });
  oldRows.sort((a, b) => b - a).forEach(row => detailSheet.deleteRow(row));
  const detailRows = ctx.questions.map(q => {
    const answer = String(answers && answers[String(q.No)] || '');
    const type = String(q.Tipe || '').toUpperCase();
    const key = String(q.Kunci || '');
    let correct = '';
    let score = '';
    if (type === 'PG') {
      correct = answer.trim() && key.trim() && answer.toUpperCase() === key.trim().toUpperCase() ? 'Ya' : 'Tidak';
      score = correct === 'Ya' ? (Number(q.Bobot) || 0) : 0;
    } else if (type === 'ISIAN' || type === 'ESSAY') {
      correct = '';
      score = '';
    } else {
      throw new Error('Tipe soal tidak valid: ' + type);
    }
    return [new Date(), className, Number(number), String(ctx.participant.Nama), q.No, type,
      String(q.Pertanyaan || ''), answer, key, correct, score, Number(q.Bobot) || 0];
  });
  if (detailRows.length) detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, ANSWER_HEADERS.length).setValues(detailRows);
  upsertWideAnswerRow_(answerSheet, ctx, answers);
  ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Status')).setValue('Selesai');
  ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Selesai')).setValue(new Date());
  ctx.sheet.getRange(ctx.rowNumber, column_(ctx.headers, 'Jawaban Sementara')).setValue(JSON.stringify(answers || {}));
  applyParticipantFormulas_(ctx.sheet, ctx.rowNumber, ctx.code);
  SpreadsheetApp.flush();
  ctx.participant.Status = 'Selesai';
}

function resultFor_(exam, code, participant) {
  const ss = getSpreadsheet_();
  const pSheet = ss.getSheetByName('Peserta_' + code);
  const pTable = table_(pSheet);
  const participantRow = pTable.rows.find(row =>
    String(row.Kelas) === String(participant.Kelas) && normalizeNumber_(row['No Absen']) === normalizeNumber_(participant['No Absen']));
  const rowNumber = participantRow ? participantRow.__rowNumber : 0;
  const detailSheet = ss.getSheetByName(detailAnswerSheetName_(code));
  const answers = table_(detailSheet).rows.filter(row =>
    String(row.Kelas) === String(participant.Kelas) && normalizeNumber_(row['No Absen']) === normalizeNumber_(participant['No Absen']));
  const showSetting = String(exam['Tampilkan Nilai'] || 'Tidak');
  const closeAt = asMillis_(exam.Tutup);
  const showScore = showSetting === 'Langsung' || (showSetting === 'Setelah ditutup' && closeAt > 0 && Date.now() >= closeAt);
  const questions = getQuestions_(code);
  const manualPending = answers.some(row =>
    String(row.Tipe).toUpperCase() !== 'PG' && row.Skor === '' && String(row.Jawaban || '').trim());
  let score = null;
  if (showScore && rowNumber >= 2) {
    applyParticipantFormulas_(pSheet, rowNumber, code);
    SpreadsheetApp.flush();
    score = scoreSummary_(answers, questions);
  }
  const explanation = String(exam['Tampilkan Pembahasan']) === 'Ya'
    ? answers.map(row => {
      const q = questions.find(item => String(item.No) === String(row.No));
      const manual = String(row.Tipe).toUpperCase() !== 'PG';
      const pending = manual && (row.Skor === '' || row.Skor == null) && String(row.Jawaban || '').trim() !== '';
      return { number: row.No, type: row.Tipe, question: row.Pertanyaan, answer: row.Jawaban,
        key: row.Kunci, correct: row.Benar === 'Ya', manual, pending, score: row.Skor, maxScore: row.Bobot,
        explanation: q ? String(q.Pembahasan || '') : '' };
    }) : [];
  return { showScore, score, manualPending, explanations: explanation };
}

function publicQuestions_(questions, data) {
  const order = Array.isArray(data.questionOrder) ? data.questionOrder : questions.map((_, i) => i);
  return order.map(index => {
    const q = questions[index];
    if (!q) throw new Error('Urutan soal tersimpan tidak valid.');
    const original = ['A', 'B', 'C', 'D', 'E'].filter(key => String(q[key] || '').trim());
    const arranged = data.optionOrder && data.optionOrder[String(index)] || original;
    return { number: q.No, type: String(q.Tipe || '').toUpperCase(), question: String(q.Pertanyaan || ''),
      image: driveThumbnail_(q.Gambar), options: arranged.map((key, i) => ({ label: String.fromCharCode(65 + i), original: key, text: String(q[key] || '') })) };
  });
}

// Menggunakan cache singkat untuk data yang sering dibaca saat login.
function cachedTable_(key, sheetName) {
  const cache = CacheService.getScriptCache();
  const value = cache.get(key);
  if (value) return JSON.parse(value);
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');
  const rows = table_(sheet).rows;
  const serialized = JSON.stringify(rows);
  if (serialized.length < 95000) cache.put(key, serialized, CACHE_SECONDS);
  return rows;
}

function getQuestions_(code) {
  const cache = CacheService.getScriptCache();
  const key = 'web:questions:' + code;
  const value = cache.get(key);
  if (value) return JSON.parse(value);
  const sheet = getSpreadsheet_().getSheetByName('Soal_' + code);
  if (!sheet) throw new Error('Sheet soal ujian tidak ditemukan.');
  const questions = table_(sheet).rows.filter(row => row.No !== '' && row.No != null);
  if (!questions.length) throw new Error('Belum ada soal untuk ujian ini. Hubungi guru.');
  const seenNumbers = {};
  questions.forEach(question => {
    const number = Number(question.No);
    const type = String(question.Tipe || '').trim().toUpperCase();
    const weightText = String(question.Bobot == null ? '' : question.Bobot).trim();
    const weight = weightText === '' ? 0 : Number(weightText);
    if (!Number.isInteger(number) || number < 1 || seenNumbers[String(number)]) {
      throw new Error('Nomor soal harus berupa angka positif dan tidak boleh berulang.');
    }
    if (!['PG', 'ISIAN', 'ESSAY'].includes(type)) throw new Error('Tipe soal harus PG, ISIAN, atau ESSAY.');
    if (!String(question.Pertanyaan || '').trim()) throw new Error('Teks pertanyaan tidak boleh kosong.');
    if (!isFinite(weight) || weight < 0) throw new Error('Bobot soal harus berupa angka nol atau lebih.');
    question.No = number;
    question.Tipe = type;
    question.Bobot = weight;
    seenNumbers[String(number)] = true;
    if (type === 'PG') {
      const options = ['A', 'B', 'C', 'D', 'E'].filter(key => String(question[key] || '').trim());
      const key = String(question.Kunci || '').trim().toUpperCase();
      if (!options.length || !options.includes(key)) throw new Error('Soal pilihan ganda harus memiliki opsi dan kunci yang sesuai.');
    }
  });
  const serialized = JSON.stringify(questions);
  if (serialized.length < 95000) cache.put(key, serialized, CACHE_SECONDS);
  return questions;
}

function table_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map(value => String(value).trim());
  const rows = [];
  values.slice(1).forEach((row, index) => {
    if (row.some(cell => cell !== '')) {
      const item = rowObject_(headers, row);
      item.__rowNumber = index + 2;
      rows.push(item);
    }
  });
  return { headers, rows };
}

function rowObject_(headers, row) {
  const result = {};
  headers.forEach((header, index) => { if (header) result[header] = row[index]; });
  return result;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  return sheet;
}

function getSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const id = properties.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Spreadsheet belum disiapkan. Jalankan menu "Siapkan spreadsheet".');
  properties.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function ensureQuestionSheet_(ss, code, addSamples) {
  const sheet = ensureSheet_(ss, 'Soal_' + code, QUESTION_HEADERS);
  sheet.getRange('E:J').setNumberFormat('@');
  sheet.getRange('D:D').setNumberFormat('@');
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['PG', 'ISIAN', 'ESSAY'], true).setAllowInvalid(false).build();
  sheet.getRange('B2:B').setDataValidation(typeRule);
  if (addSamples && sheet.getLastRow() < 2) {
    const samples = [
      [1, 'PG', 'Hitung nilai $\\frac{1}{2}+\\frac{1}{4}$.', '', '$\\frac{1}{2}$', '$\\frac{3}{4}$', '$1$', '$\\frac{1}{4}$', '', 'B', 2, 'Samakan penyebut menjadi 4: $\\frac{2}{4}+\\frac{1}{4}=\\frac{3}{4}$.'],
      [2, 'PG', 'Jika $2x+3=11$, nilai $x$ adalah …', '', '2', '3', '4', '5', '', 'C', 2, 'Kurangi 3 dari kedua sisi, lalu bagi 2.'],
      [3, 'PG', 'Sederhanakan $3a+2a-4$.', '', '$5a-4$', '$a-4$', '$5a+4$', '$6a$', '', 'A', 2, 'Gabungkan suku sejenis: $3a+2a=5a$.'],
      [4, 'ISIAN', 'Hitung $5^2-3^2$.', '', '', '', '', '', '', '16', 2, 'Karena $25-9=16$.'],
      [5, 'ISIAN', 'Tuliskan $0{,}75$ sebagai pecahan.', '', '', '', '', '', '', '3/4|0,75', 2, '$0{,}75=\\frac{75}{100}=\\frac{3}{4}$.'],
      [6, 'ESSAY', 'Sebuah persegi panjang panjangnya $8$ cm dan lebarnya $5$ cm. Hitung keliling dan luasnya.', '', '', '', '', '', '', '', 4, 'Keliling $=2(p+l)=26$ cm dan luas $=p\\times l=40$ cm².']
    ];
    sheet.getRange(2, 1, samples.length, QUESTION_HEADERS.length).setValues(samples);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, QUESTION_HEADERS.length);
  return sheet;
}

function ensureParticipantSheets_(ss, code) {
  const participants = ensureSheet_(ss, 'Peserta_' + code, PARTICIPANT_HEADERS);
  migrateParticipantHeaders_(participants);
  const questionSheet = ss.getSheetByName('Soal_' + code);
  const questions = questionSheet ? table_(questionSheet).rows.filter(row => row.No !== '' && row.No != null) : [];
  const answerSheets = ensureAnswerSheets_(ss, code, questions);
  participants.hideColumns(15, 3);
  table_(participants).rows.forEach(row => {
    if (String(row.Kelas || '').trim()) applyParticipantFormulas_(participants, row.__rowNumber, code);
  });
  participants.getRange('B:B').setNumberFormat('0');
  participants.getRange('E:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  participants.getRange('J:N').setNumberFormat('0.0');
  participants.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Login', 'Mengerjakan', 'Selesai'], true).setAllowInvalid(false).build());
  participants.setFrozenRows(1);
  answerSheets.wide.setFrozenRows(1);
  answerSheets.detail.setFrozenRows(1);
  participants.autoResizeColumns(1, 14);
  answerSheets.wide.autoResizeColumns(1, answerSheets.wide.getLastColumn());
  answerSheets.detail.getRange('K2:K').setBackground('#fff2cc');
  answerSheets.detail.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['PG', 'ISIAN', 'ESSAY'], true).setAllowInvalid(false).build());
  answerSheets.detail.getRange('J2:J').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Ya', 'Tidak'], true).setAllowInvalid(false).build());
  answerSheets.detail.autoResizeColumns(1, ANSWER_HEADERS.length);
}

function detailAnswerSheetName_(code) {
  return 'Detail_' + code;
}

function ensureAnswerSheets_(ss, code, questions) {
  const wideName = 'Jawaban_' + code;
  const detailName = detailAnswerSheetName_(code);
  let wide = ss.getSheetByName(wideName);
  let detail = ss.getSheetByName(detailName);
  let migratedLegacy = false;

  if (wide && isDetailAnswerSheet_(wide)) {
    if (detail) throw new Error('Ada sheet jawaban lama dan sheet detail sekaligus. Periksa "' + detailName + '" sebelum migrasi.');
    wide.setName(detailName);
    detail = wide;
    wide = null;
    migratedLegacy = true;
  }
  if (!detail) detail = ensureSheet_(ss, detailName, ANSWER_HEADERS);
  if (!wide) wide = ss.insertSheet(wideName);
  ensureWideAnswerLayout_(wide, questions || []);
  if (migratedLegacy) backfillWideAnswerRows_(wide, detail, questions || []);
  return { wide, detail };
}

function isDetailAnswerSheet_(sheet) {
  const width = Math.max(ANSWER_HEADERS.length, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(value => String(value).trim());
  return ANSWER_HEADERS.every((header, index) => headers[index] === header);
}

function wideAnswerHeaders_(questions) {
  const headers = WIDE_ANSWER_BASE_HEADERS.slice();
  questions.slice().sort((a, b) => Number(a.No) - Number(b.No)).forEach(question => {
    const number = Number(question.No);
    headers.push('Jawaban ' + number);
    if (String(question.Tipe || '').toUpperCase() !== 'PG') headers.push('Nilai Guru ' + number);
  });
  headers.push('Skor PG', 'Nilai Akumulasi (/100)');
  return headers;
}

function ensureWideAnswerLayout_(sheet, questions) {
  const values = sheet.getDataRange().getValues();
  const currentHeaders = values.length ? values[0].map(value => String(value).trim()).filter(Boolean) : [];
  const required = wideAnswerHeaders_(questions);
  const target = required;
  if (currentHeaders.length === target.length && currentHeaders.every((header, index) => header === target[index])) {
    styleHeader_(sheet, target.length);
    applyWideAnswerFormatting_(sheet, target, questions);
    values.slice(1).forEach((row, index) => {
      if (String(row[1] || '').trim() && normalizeNumber_(row[2])) {
        applyWideAnswerTotals_(sheet, target, index + 2, codeFromSheetName_(sheet.getName()));
      }
    });
    return target;
  }

  const oldHeaders = values.length ? values[0].map(value => String(value).trim()) : [];
  const oldIndex = {};
  oldHeaders.forEach((header, index) => {
    if (!header) return;
    oldIndex[header] = index;
    const oldScore = header.match(/^Skor Essay (\d+)$/);
    if (oldScore) oldIndex['Nilai Guru ' + oldScore[1]] = index;
    if (header === 'Nilai /100' || header === 'Nilai Akumulasi /100' ||
        header === 'Nilai Akumulasi (/100)') {
      oldIndex['Nilai Akumulasi (/100)'] = index;
    }
  });
  const rows = values.slice(1)
    .filter(row => row.some(value => value !== ''))
    .map(row => target.map(header => Object.prototype.hasOwnProperty.call(oldIndex, header) ? row[oldIndex[header]] : ''));
  sheet.clearContents();
  sheet.getRange(1, 1, 1, target.length).setValues([target]);
  if (rows.length) sheet.getRange(2, 1, rows.length, target.length).setValues(rows);
  styleHeader_(sheet, target.length);
  applyWideAnswerFormatting_(sheet, target, questions);
  rows.forEach((row, index) => {
    if (String(row[1] || '').trim() && normalizeNumber_(row[2])) {
      applyWideAnswerTotals_(sheet, target, index + 2, codeFromSheetName_(sheet.getName()));
    }
  });
  return target;
}

function backfillWideAnswerRows_(wide, detail, questions) {
  const headers = wideAnswerHeaders_(questions);
  const columns = {};
  headers.forEach((header, index) => { columns[header] = index; });
  const grouped = {};
  table_(detail).rows.forEach(answer => {
    const studentKey = String(answer.Kelas).trim() + '|' + normalizeNumber_(answer['No Absen']);
    let row = grouped[studentKey];
    if (!row) {
      row = new Array(headers.length).fill('');
      row[columns['Waktu Selesai']] = answer.Waktu || '';
      row[columns.Kelas] = answer.Kelas;
      row[columns['No Absen']] = Number(normalizeNumber_(answer['No Absen']));
      row[columns.Nama] = answer.Nama;
      grouped[studentKey] = row;
    }
    const number = Number(answer.No);
    const answerColumn = columns['Jawaban ' + number];
    if (answerColumn !== undefined) row[answerColumn] = String(answer.Jawaban || '');
    if (String(answer.Tipe).toUpperCase() !== 'PG') {
      const scoreColumn = columns['Nilai Guru ' + number];
      if (scoreColumn !== undefined) row[scoreColumn] = answer.Skor;
    }
  });
  const rows = Object.keys(grouped).map(key => grouped[key]);
  if (rows.length) wide.getRange(2, 1, rows.length, headers.length).setValues(rows);
  rows.forEach((row, index) => {
    if (String(row[1] || '').trim() && normalizeNumber_(row[2])) {
      applyWideAnswerTotals_(wide, headers, index + 2, codeFromSheetName_(wide.getName()));
    }
  });
}

function codeFromSheetName_(sheetName) {
  return String(sheetName).slice('Jawaban_'.length);
}

function applyWideAnswerTotals_(sheet, headers, row, code) {
  if (row < 2 || !String(sheet.getRange(row, 2).getValue() || '').trim() ||
      !normalizeNumber_(sheet.getRange(row, 3).getValue())) return;
  const indexes = {};
  headers.forEach((header, index) => { indexes[header] = index + 1; });
  const classroom = String(sheet.getRange(row, 2).getValue()).trim();
  const number = normalizeNumber_(sheet.getRange(row, 3).getValue());
  const detail = getSpreadsheet_().getSheetByName(detailAnswerSheetName_(code));
  const answers = table_(detail).rows.filter(answer =>
    String(answer.Kelas).trim() === classroom && normalizeNumber_(answer['No Absen']) === number);
  const questions = getQuestions_(code);
  const questionByNumber = {};
  questions.forEach(question => { questionByNumber[String(Number(question.No))] = question; });
  const pgScore = answers.reduce((total, answer) =>
    String(answer.Tipe).toUpperCase() === 'PG' ? total + numericScore_(answer.Skor) : total, 0);
  let manualScore = 0;
  headers.forEach((header, index) => {
    const manualScoreHeader = header.match(/^Nilai Guru (\d+)$/);
    if (!manualScoreHeader) return;
    const question = questionByNumber[manualScoreHeader[1]];
    if (!question || String(question.Tipe).toUpperCase() === 'PG') return;
    const scoreColumn = index + 1;
    const scoreCell = sheet.getRange(row, scoreColumn);
    const value = scoreCell.getValue();
    if (value == null || String(value).trim() === '') {
      scoreCell.setBackground('#fff2cc');
      return;
    }
    const score = Number(value);
    const maximum = Number(question.Bobot) || 0;
    if (isFinite(score) && score >= 0 && score <= maximum) {
      manualScore += score;
      scoreCell.setBackground('#fff2cc');
    } else {
      scoreCell.setBackground('#f4cccc');
    }
  });
  const maximum = questions.reduce((total, question) => total + (Number(question.Bobot) || 0), 0);
  const total = pgScore + manualScore;
  const value = maximum > 0 ? Math.round(total / maximum * 1000) / 10 : 0;
  sheet.getRange(row, indexes['Skor PG']).setValue(pgScore);
  const formulaTerms = [columnLetter_(indexes['Skor PG']) + row];
  questions.forEach(question => {
    if (String(question.Tipe).toUpperCase() === 'PG') return;
    const scoreColumn = indexes['Nilai Guru ' + Number(question.No)];
    if (!scoreColumn) return;
    formulaTerms.push(columnLetter_(scoreColumn) + row);
  });
  const weightTerms = questions.map(question =>
    'Soal_' + code + '!$K$' + question.__rowNumber);
  const accumulationCell = sheet.getRange(row, indexes['Nilai Akumulasi (/100)']);
  if (maximum > 0) {
    accumulationCell.setFormula(
      '=(' + formulaTerms.join('+') + ')/(' + weightTerms.join('+') + ')*100'
    );
  } else {
    accumulationCell.setValue(0);
  }
  syncWideManualScoresToDetail_(sheet, headers, row, code, answers, questions);
  return { automatic: pgScore, essay: manualScore, total, max: maximum, value };
}

function columnLetter_(column) {
  let value = Number(column);
  let letters = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function syncWideManualScoresToDetail_(sheet, headers, row, code, answers, questions) {
  const details = {};
  answers.forEach(answer => {
    if (String(answer.Tipe).toUpperCase() !== 'PG') {
      details[String(Number(answer.No))] = answer.__rowNumber;
    }
  });
  const indexes = {};
  headers.forEach((header, index) => { indexes[header] = index + 1; });
  const detail = getSpreadsheet_().getSheetByName(detailAnswerSheetName_(code));
  headers.forEach((header, index) => {
    const manualScoreHeader = header.match(/^Nilai Guru (\d+)$/);
    if (!manualScoreHeader) return;
    const question = questions.find(item => Number(item.No) === Number(manualScoreHeader[1]));
    if (!question || String(question.Tipe).toUpperCase() === 'PG') return;
    const detailRow = details[manualScoreHeader[1]];
    const scoreColumn = index + 1;
    if (!detailRow || !scoreColumn) return;
    const value = sheet.getRange(row, scoreColumn).getValue();
    const maximum = Number(question.Bobot) || 0;
    if (value !== '' && value != null && isFinite(Number(value)) && Number(value) >= 0 && Number(value) <= maximum) {
      detail.getRange(detailRow, ANSWER_HEADERS.indexOf('Skor') + 1).setValue(Number(value));
    }
  });
}

function applyWideAnswerFormatting_(sheet, headers, questions) {
  sheet.setFrozenRows(1);
  sheet.getRange(2, 3, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('0');
  const questionByNumber = {};
  questions.forEach(question => { questionByNumber[String(Number(question.No))] = question; });
  headers.forEach((header, index) => {
    const column = index + 1;
    const manualScore = header.match(/^Nilai Guru (\d+)$/);
    if (manualScore) {
      sheet.getRange(2, column, Math.max(1, sheet.getMaxRows() - 1), 1)
        .setBackground('#fff2cc')
        .setNumberFormat('0.0');
      const question = questionByNumber[manualScore[1]];
      if (question) {
        const validation = SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, Number(question.Bobot) || 0)
          .setHelpText('Nilai harus dari 0 sampai bobot maksimal soal (' + (Number(question.Bobot) || 0) + ').')
          .setAllowInvalid(false)
          .build();
        sheet.getRange(2, column, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(validation);
      }
    } else if (/^Jawaban \d+$/.test(header)) {
      sheet.getRange(2, column, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
    }
  });
}

function upsertWideAnswerRow_(sheet, ctx, answers) {
  const headers = ensureWideAnswerLayout_(sheet, ctx.questions);
  const table = table_(sheet);
  const existing = table.rows.find(row =>
    String(row.Kelas) === String(ctx.participant.Kelas) &&
    normalizeNumber_(row['No Absen']) === normalizeNumber_(ctx.participant['No Absen']));
  const rowNumber = existing ? existing.__rowNumber : firstEmptyParticipantRow_(sheet);
  const oldValues = existing ? sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0] : new Array(headers.length).fill('');
  const values = oldValues.slice();
  const columnByHeader = {};
  headers.forEach((header, index) => { columnByHeader[header] = index; });
  values[columnByHeader['Waktu Selesai']] = new Date();
  values[columnByHeader.Kelas] = ctx.participant.Kelas;
  values[columnByHeader['No Absen']] = Number(normalizeNumber_(ctx.participant['No Absen']));
  values[columnByHeader.Nama] = ctx.participant.Nama;
  ctx.questions.forEach(question => {
    const number = Number(question.No);
    const answerIndex = columnByHeader['Jawaban ' + number];
    if (answerIndex !== undefined) values[answerIndex] = String(answers && answers[String(question.No)] || '');
    if (String(question.Tipe || '').toUpperCase() !== 'PG') {
      const scoreIndex = columnByHeader['Nilai Guru ' + number];
      if (scoreIndex !== undefined) values[scoreIndex] = '';
    }
  });
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
  applyWideAnswerTotals_(sheet, headers, rowNumber, ctx.code);
}

function syncManualScores_(event) {
  const sheet = event.range.getSheet();
  const code = sheet.getName().slice('Jawaban_'.length);
  const detail = getSpreadsheet_().getSheetByName(detailAnswerSheetName_(code));
  if (!detail || event.range.getLastRow() < 2) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(value => String(value).trim());
  const identity = headerIndex_(headers, 'Kelas', 'No Absen');
  const firstRow = Math.max(2, event.range.getRow());
  const lastRow = event.range.getLastRow();
  const firstColumn = event.range.getColumn();
  const lastColumn = event.range.getLastColumn();
  const edited = sheet.getRange(firstRow, firstColumn, lastRow - firstRow + 1, lastColumn - firstColumn + 1).getValues();
  const details = table_(detail);
  const detailByAnswer = {};
  details.rows.forEach(row => {
    if (String(row.Tipe).toUpperCase() !== 'PG') {
      detailByAnswer[answerRecordKey_(row.Kelas, row['No Absen'], row.No)] = row.__rowNumber;
    }
  });
  edited.forEach((row, rowOffset) => {
    row.forEach((value, columnOffset) => {
      const header = headers[firstColumn + columnOffset - 1];
      const scoreColumn = header.match(/^Nilai Guru (\d+)$/);
      if (!scoreColumn) return;
      const sourceRow = firstRow + rowOffset;
      const rowValues = sheet.getRange(sourceRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const classroom = rowValues[identity.classroom];
      const number = rowValues[identity.number];
      const detailRow = detailByAnswer[answerRecordKey_(classroom, number, scoreColumn[1])];
      if (!detailRow) return;
      const maxScore = Number(detail.getRange(detailRow, ANSWER_HEADERS.indexOf('Bobot') + 1).getValue()) || 0;
      if (value !== '' && value != null &&
          (typeof value === 'boolean' || String(value).trim() === '' || !isFinite(Number(value)) ||
            Number(value) < 0 || Number(value) > maxScore)) {
        const scoreCell = sheet.getRange(sourceRow, firstColumn + columnOffset);
        if (event.range.getNumRows() === 1 && event.range.getNumColumns() === 1 && event.oldValue !== undefined) {
          scoreCell.setValue(event.oldValue);
        } else {
          scoreCell.clearContent();
        }
        sheet.getParent().toast('Nilai harus berupa angka dari 0 sampai ' + maxScore + '.', 'Nilai guru tidak valid', 5);
        return;
      }
      detail.getRange(detailRow, ANSWER_HEADERS.indexOf('Skor') + 1).setValue(value);
      applyWideAnswerTotals_(sheet, headers, sourceRow, code);
      const participantSheet = getSpreadsheet_().getSheetByName('Peserta_' + code);
      if (participantSheet) {
        const participant = table_(participantSheet).rows.find(item =>
          String(item.Kelas).trim() === String(classroom).trim() &&
          normalizeNumber_(item['No Absen']) === normalizeNumber_(number));
        if (participant) applyParticipantFormulas_(participantSheet, participant.__rowNumber, code);
      }
    });
  });
}

function headerIndex_(headers, classroomHeader, numberHeader) {
  const classroom = headers.indexOf(classroomHeader);
  const number = headers.indexOf(numberHeader);
  if (classroom < 0 || number < 0) throw new Error('Kolom identitas tidak ditemukan pada sheet jawaban.');
  return { classroom, number };
}

function answerRecordKey_(classroom, number, questionNumber) {
  return String(classroom).trim() + '|' + normalizeNumber_(number) + '|' + String(Number(questionNumber));
}

function numericScore_(value) {
  if (value == null || String(value).trim() === '') return 0;
  const number = Number(value);
  if (!isFinite(number)) throw new Error('Skor atau bobot soal harus berupa angka.');
  return number;
}

function scoreSummary_(answers, questions) {
  const automatic = answers.reduce((sum, answer) =>
    String(answer.Tipe).toUpperCase() === 'PG' ? sum + numericScore_(answer.Skor) : sum, 0);
  const essay = answers.reduce((sum, answer) =>
    String(answer.Tipe).toUpperCase() !== 'PG' ? sum + numericScore_(answer.Skor) : sum, 0);
  const total = automatic + essay;
  const max = questions.reduce((sum, question) => sum + numericScore_(question.Bobot), 0);
  const objective = answers.filter(answer => String(answer.Tipe).toUpperCase() === 'PG');
  const correct = objective.filter(answer => String(answer.Benar).trim() === 'Ya').length;
  return { automatic, essay, total, max, value: max > 0 ? Math.round(total / max * 1000) / 10 : 0,
    correct, questionCount: questions.length, objectiveCount: objective.length };
}

function applyParticipantFormulas_(sheet, row, code) {
  const classroom = String(sheet.getRange(row, 1).getValue() || '').trim();
  const number = normalizeNumber_(sheet.getRange(row, 2).getValue());
  if (!classroom || !number) return;
  const answerSheet = getSpreadsheet_().getSheetByName('Jawaban_' + code);
  const answerTable = table_(answerSheet);
  const answerRow = answerTable.rows.find(answer =>
    String(answer.Kelas).trim() === classroom && normalizeNumber_(answer['No Absen']) === number);
  const score = answerRow
    ? applyWideAnswerTotals_(answerSheet, answerTable.headers, answerRow.__rowNumber, code)
    : { automatic: 0, essay: 0, total: 0, max: getQuestions_(code).reduce((sum, question) => sum + (Number(question.Bobot) || 0), 0), value: 0 };
  sheet.getRange(row, 10, 1, 5).setValues([[
    score.automatic, score.essay, score.total, score.max, score.value
  ]]);
}

function firstEmptyParticipantRow_(sheet) {
  const maxRows = sheet.getMaxRows();
  const values = maxRows > 1 ? sheet.getRange(2, 1, maxRows - 1, 1).getValues() : [];
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === '') return i + 2;
  }
  sheet.insertRowAfter(maxRows);
  return maxRows + 1;
}

function applyBaseFormatting_(students, exams) {
  students.getRange('B:B').setNumberFormat('0');
  applyExamValidation_(exams);
  exams.getRange('E:F').setNumberFormat('yyyy-mm-dd hh:mm');
  exams.getRange('G:G').setNumberFormat('0');
  exams.setFrozenRows(1);
  students.setFrozenRows(1);
  students.autoResizeColumns(1, STUDENT_HEADERS.length);
  exams.autoResizeColumns(1, EXAM_HEADERS.length);
}

function applyExamValidation_(sheet) {
  const setList = (column, values) => sheet.getRange(column + '2:' + column).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build());
  setList('H', ['Tidak', 'Langsung', 'Setelah ditutup']);
  setList('I', ['Ya', 'Tidak']);
  ['L', 'M'].forEach(column => setList(column, ['Ya', 'Tidak']));
  setList('N', ['Aktif', 'Nonaktif', 'Diarsipkan']);
}

function styleHeader_(sheet, columnCount) {
  sheet.getRange(1, 1, 1, columnCount).setBackground('#17365d').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.setFrozenRows(1);
}

function classAllowed_(value, classroom) {
  const classes = String(value || '').split(',').map(item => item.trim().toUpperCase());
  return classes.includes('SEMUA') || classes.includes(String(classroom).trim().toUpperCase());
}

function normalizeNumber_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!/^\d+$/.test(text)) return '';
  return String(Number(text));
}

function asMillis_(value) {
  if (value == null || String(value).trim() === '') return 0;
  if (value instanceof Date) {
    const time = value.getTime();
    if (!isNaN(time)) return time;
    throw new Error('Tanggal ujian tidak valid. Periksa kolom Buka dan Tutup.');
  }
  const parsed = new Date(value).getTime();
  if (isNaN(parsed)) throw new Error('Tanggal ujian tidak valid. Periksa kolom Buka dan Tutup.');
  return parsed;
}

function pauseSeconds_(exam) {
  const value = exam['Jeda (detik)'];
  if (value == null || String(value).trim() === '') return 60;
  const seconds = Number(value);
  if (!isFinite(seconds) || seconds < 0) throw new Error('Jeda harus berupa angka nol atau lebih.');
  return seconds;
}

function durationMinutes_(exam) {
  const minutes = Number(exam['Durasi (menit)']);
  if (!isFinite(minutes) || minutes <= 0) throw new Error('Durasi ujian harus berupa angka lebih dari nol menit.');
  return minutes;
}

function maxViolations_(exam) {
  const value = exam['Maks Pelanggaran'];
  if (value == null || String(value).trim() === '') return 0;
  const maximum = Number(value);
  if (!Number.isInteger(maximum) || maximum < 0) throw new Error('Maks Pelanggaran harus berupa angka bulat nol atau lebih.');
  return maximum;
}

function isYes_(value) { return String(value || '').trim().toLowerCase() === 'ya'; }

function parseJson_(value, fallback) {
  if (value == null || String(value).trim() === '') return fallback;
  try { return JSON.parse(String(value)); }
  catch (error) { throw new Error('Data sesi tidak dapat dibaca. Hubungi guru untuk bantuan.'); }
}

function answerMatches_(answer, keys) {
  const normalized = normalizeAnswer_(answer);
  if (!normalized) return false;
  return String(keys || '').split('|').some(key => normalizeAnswer_(key) === normalized);
}

function normalizeAnswer_(value) {
  const text = String(value || '').replace(/\s+/g, '').replace(/,/g, '.').toLowerCase();
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(text)) return String(Number(text));
  const fraction = text.match(/^([+-]?\d+)\/([+-]?\d+)$/);
  if (fraction && Number(fraction[2]) !== 0) return String(Number(fraction[1]) / Number(fraction[2]));
  return text;
}

function driveThumbnail_(value) {
  const url = String(value || '').trim();
  if (!url) return '';
  const match = url.match(/(?:\/d\/|[?&]id=)([a-zA-Z0-9_-]+)/);
  return match ? 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w1000' : url;
}

function shuffle_(items) {
  for (let i = items.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [items[i], items[j]] = [items[j], items[i]];
  }
}

function randomToken_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let token = '';
  for (let i = 0; i < 6; i++) token += chars.charAt(Math.floor(Math.random() * chars.length));
  return token;
}

function uniqueToken_() {
  const used = cachedTable_('web:exams', SHEET_EXAMS)
    .map(row => String(row.Token || '').trim());
  let token = randomToken_();
  while (used.includes(token)) token = randomToken_();
  return token;
}

function examByCode_(code) {
  return cachedTable_('web:exams', SHEET_EXAMS).find(row => String(row.Kode || '').trim().toUpperCase() === code);
}

function writeExamField_(exam, header, value) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_EXAMS);
  const t = table_(sheet);
  const row = t.rows.find(item => String(item.Kode).trim().toUpperCase() === String(exam.Kode).trim().toUpperCase());
  if (row) sheet.getRange(row.__rowNumber, column_(t.headers, header)).setValue(value);
}

function column_(headers, name) {
  const index = headers.indexOf(name);
  if (index < 0) throw new Error('Kolom "' + name + '" tidak ditemukan.');
  return index + 1;
}

function appError_(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function uniqueSheetName_(ss, base) {
  let name = base;
  let suffix = 2;
  while (ss.getSheetByName(name)) name = base + '_' + suffix++;
  return name;
}
