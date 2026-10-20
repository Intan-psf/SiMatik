<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store');

function respond(int $status, array $body): never
{
    http_response_code($status);
    echo json_encode($body, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Allow: POST');
    respond(405, ['ok' => false, 'error' => 'Metode permintaan tidak didukung.']);
}

if (!extension_loaded('curl')) {
    respond(500, ['ok' => false, 'error' => 'PHP cURL belum aktif. Aktifkan ekstensi cURL pada PHP.']);
}

$rawBody = file_get_contents('php://input');
if ($rawBody === false || strlen($rawBody) > 100000) {
    respond(413, ['ok' => false, 'error' => 'Ukuran permintaan terlalu besar.']);
}

$payload = json_decode($rawBody, true);
$allowedActions = [
    'getLoginOptions',
    'loginStudent',
    'startExam',
    'saveAnswers',
    'reportViolation',
    'submitExam',
];
if (!is_array($payload) || !isset($payload['action']) || !in_array($payload['action'], $allowedActions, true)) {
    respond(400, ['ok' => false, 'error' => 'Permintaan tidak valid.']);
}
if (isset($payload['args']) && !is_array($payload['args'])) {
    respond(400, ['ok' => false, 'error' => 'Format data permintaan tidak valid.']);
}

require __DIR__ . '/config.php';
$url = filter_var($appsScriptWebAppUrl ?? '', FILTER_VALIDATE_URL);
$parts = $url ? parse_url($url) : false;
if (
    !$parts ||
    ($parts['scheme'] ?? '') !== 'https' ||
    ($parts['host'] ?? '') !== 'script.google.com' ||
    !preg_match('~^/macros/s/[^/]+/exec$~', $parts['path'] ?? '') ||
    str_contains($parts['path'] ?? '', 'PASTE_DEPLOYMENT_ID')
) {
    respond(500, ['ok' => false, 'error' => 'Isi URL deployment Apps Script yang benar di config.php.']);
}

$curl = curl_init($url);
curl_setopt_array($curl, [
    CURLOPT_POST => true,
    CURLOPT_POSTFIELDS => json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES),
    CURLOPT_HTTPHEADER => ['Content-Type: application/json; charset=utf-8'],
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_FOLLOWLOCATION => true,
    CURLOPT_MAXREDIRS => 5,
    CURLOPT_CONNECTTIMEOUT => 10,
    CURLOPT_TIMEOUT => 90,
    CURLOPT_PROTOCOLS => CURLPROTO_HTTPS,
    CURLOPT_REDIR_PROTOCOLS => CURLPROTO_HTTPS,
]);

$responseBody = curl_exec($curl);
$curlError = curl_error($curl);
$httpStatus = (int) curl_getinfo($curl, CURLINFO_RESPONSE_CODE);
curl_close($curl);

if ($responseBody === false) {
    error_log('Apps Script relay failed: ' . $curlError);
    respond(502, ['ok' => false, 'error' => 'Tidak dapat menghubungi Apps Script. Periksa internet dan URL deployment.']);
}

$response = json_decode($responseBody, true);
if ($httpStatus < 200 || $httpStatus >= 300 || !is_array($response) || !array_key_exists('ok', $response)) {
    error_log('Unexpected Apps Script response; HTTP ' . $httpStatus);
    respond(502, ['ok' => false, 'error' => 'Respons Apps Script tidak sesuai. Pastikan versi deployment sudah diperbarui.']);
}

respond(200, $response);
