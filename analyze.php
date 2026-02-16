<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

//Load Excel
if (!isset($_FILES['excel']) || $_FILES['excel']['error'] !== UPLOAD_ERR_OK) {
    die("Σφάλμα στο ανέβασμα του αρχείου, παρακαλώ σιγουρευτείτε ότι ανεβάσατε κατάλληλο αρχείο Excel.");
}

$tmpFilePath = $_FILES['excel']['tmp_name'];
$spreadsheet = IOFactory::load($tmpFilePath);
$sheet = $spreadsheet->getActiveSheet();

$subjects = [];
$BSubjects = [];
$ASubjects = [];
$ASubjectsUnpassed = [];

for ($row = 3; ; $row++) {
    $code = trim($sheet->getCell(Coordinate::stringFromColumnIndex(1) . $row)->getValue());

    $name = trim($sheet->getCell(Coordinate::stringFromColumnIndex(2) . $row)->getValue());
    if ($name === '') break;

    $mark = $sheet->getCell(Coordinate::stringFromColumnIndex(3) . $row)->getValue();
    if ($mark > 10) $mark /= 10;

    $ects = $sheet->getCell(Coordinate::stringFromColumnIndex(12) . $row)->getValue();
    $BTag = $sheet->getCell(Coordinate::stringFromColumnIndex(17) . $row)->getValue();
    $ATag = $sheet->getCell(Coordinate::stringFromColumnIndex(14) . $row)->getValue();

    if (!is_numeric($mark) || !is_numeric($ects)) continue;

    $isitA = (mb_stripos((string)$ATag, 'ΥΠΟΧΡΕΩΤΙΚΟ') !== false 
           || mb_stripos((string)$ATag, 'ΞΕΝΗ ΓΛΩΣΣΑ') !== false);

    if ($isitA && $mark < 5) {
        $ASubjectsUnpassed[] = $name;
        continue;
    }

    if ($mark < 5) continue;

    $index = count($subjects);
    $subjects[] = [
        'code' => $code,
        'name' => $name,
        'ects' => (int)$ects,
        'mark' => (float)$mark,
    ];

    if (mb_stripos((string)$BTag, 'ΠΙΝΑΚΑΣ Β') !== false) {
        $BSubjects[] = $index;
    }

    if ($isitA) {
        $ASubjects[] = $index;
    }
}

//Validating

$minECTS = 240;
$maxECTS = 242;
if ($ASubjectsUnpassed) {
    echo "<strong>Error:</strong> Τα ακόλουθα υποχρεωτικά μαθήματα δεν περάστηκαν:<br>";
    foreach ($ASubjectsUnpassed as $unpassed) {
        echo " - " . htmlspecialchars($unpassed) . "<br>";
    }
    exit(1);
}

$minBSubjects = 4;
if (count($BSubjects) < $minBSubjects) {
    echo "<strong>Error:</strong> Βρέθηκαν μόνο " . count($BSubjects) . " μαθήματα Β κατηγορίας: χρειάζονται τουλάχιστον $minBSubjects.<br>";
    exit(1);
}

function generateLP($subjects, $BSubjects, $ASubjects, $minECTS, $maxECTS) {
    $lines = [];

    $obj = [];
    foreach ($subjects as $i => $s) {
        $obj[] = ($s['mark'] * $s['ects']) . " m" . ($i + 3);
    }
    $lines[] = "max: " . implode(' + ', $obj) . ";\n";

    $denom = [];
    foreach ($subjects as $i => $s) {
        $denom[] = $s['ects'] . " m" . ($i + 3);
    }
    $lines[] = "denom: " . implode(' + ', $denom) . " = 1;\n";

    $lines[] = "ects1: -" . $minECTS . " scale + 1 >= 0 ;\n";
    $lines[] = "ects2: -" . $maxECTS . " scale + 1 <= 0 ;\n";

    if ($BSubjects) {
        $minBSubjects = 4;
        $terms = [];
        foreach ($BSubjects as $i) {
            $terms[] = $subjects[$i]['ects'] . " m" . ($i + 3);
        }
        $lines[] = "typeB: -" . ($minBSubjects * 5) . " scale + " . implode(' + ', $terms) . " >= 0 ;\n";
    }

    foreach ($ASubjects as $i) {
        $lines[] = "assign" . ($i + 3) . ": m" . ($i + 3) . " = scale;";
    }

    $zVars = [];
    foreach ($subjects as $i => $s) {
        if (in_array($i, $ASubjects)) continue;
        $m = "m" . ($i + 3);
        $z = "z" . ($i + 3);
        $lines[] = "$m <= 10 $z ;";
        $lines[] = "$m - scale - 10 $z >= -10 ;";
        $lines[] = "$m - scale + 10 $z <= 10 ;";
        $zVars[] = $z;
    }

    $lines[] = "bin " . implode(', ', $zVars) . " ;\n";

    return implode("\n", $lines);
}

//Lp Solve
file_put_contents('model.lp', generateLP($subjects, $BSubjects, $ASubjects, 240, 242));
exec("lp_solve model.lp", $output, $status);

if ($status === 2) {
    file_put_contents('model.lp', generateLP($subjects, $BSubjects, $ASubjects, 240, 300));
    exec("lp_solve model.lp", $output, $status);
}

if ($status === 2) {
    echo "<strong>Αδυναμία λύσης:</strong> Δεν υπάρχει συνδυασμός που να ικανοποιεί όλους τους περιορισμούς.<br>";
    exit(1);
}
if ($status !== 0) {
    echo "<strong>Το lp_solve απέτυχε.</strong><br>";
    exit(1);
}

//Output
$result = [];
$parsing = false;

foreach ($output as $line) {
    if (!$parsing && str_contains($line, 'Actual values of the variables:')) {
        $parsing = true;
        continue;
    }
    if ($parsing) {
        $line = trim($line);
        if ($line === '') break;
        if (preg_match('/^m(\d+)\s+([\d.]+)/', $line, $m)) {
            $result[(int)$m[1]] = (float)$m[2];
        }
    }
}

if (!$result) {
    echo "<strong>Error:</strong> Δεν βρέθηκε λύση στο lp_solve output.<br>";
    exit(1);
}

//Final report
ob_start();

echo "<h2>Επιλεγμένα Μαθήματα:</h2>";
$totalECTS = $numerator = 0;
$selected = [];

foreach ($subjects as $i => $s) {
    $mVar = $i + 3;
    if (!empty($result[$mVar])) {
        $selected[] = $i;
        $totalECTS += $s['ects'];
        $numerator += $s['mark'] * $s['ects'];
        echo "<div class='subject-line'><div class='left'>(" . htmlspecialchars($s['code']) . ") " . htmlspecialchars($s['name']) . "</div><div class='right'>ECTS: {$s['ects']}, Βαθμός: " . number_format($s['mark'], 1) . "</div></div>";

    }
}

$average = $numerator / $totalECTS;
echo "<div class='summary'><strong>Συνολικά ECTS:</strong> $totalECTS</div>";
echo "<div class='summary'><strong>Βαθμός:</strong> " . round($average, 2) . "</div>";

//Excluded Subjects
echo "<hr class='divider'>";
echo "<h2>Μαθήματα που δεν χρησιμοποιήθηκαν:</h2>";
$hasExcluded = false;
foreach ($subjects as $i => $s) {
    if (!in_array($i, $selected)) {
        $hasExcluded = true;
        echo "<div class='subject-line excluded'><div class='left'>(" . htmlspecialchars($s['code']) . ") " . htmlspecialchars($s['name']) . "</div><div class='right'>ECTS: {$s['ects']}, Mark: " . number_format($s['mark'], 1) . "</div></div>";

    }
}
if (!$hasExcluded) {
    echo "<div class='subject-line excluded'>Όλα τα μαθήματα χρησιμοποιήθηκαν.</div>";
}

function generateLPMin($subjects, $BSubjects, $ASubjects, $minECTS, $maxECTS) {
    $lines = [];

    $obj = [];
    foreach ($subjects as $i => $s) {
        $obj[] = ($s['mark'] * $s['ects']) . " m" . ($i + 3);
    }
    $lines[] = "min: " . implode(' + ', $obj) . ";\n";

    $denom = [];
    foreach ($subjects as $i => $s) {
        $denom[] = $s['ects'] . " m" . ($i + 3);
    }
    $lines[] = "denom: " . implode(' + ', $denom) . " = 1;\n";

    $lines[] = "ects1: -" . $minECTS . " scale + 1 >= 0 ;\n";
    $lines[] = "ects2: -" . $maxECTS . " scale + 1 <= 0 ;\n";

    if ($BSubjects) {
        $minBSubjects = 4;
        $terms = [];
        foreach ($BSubjects as $i) {
            $terms[] = $subjects[$i]['ects'] . " m" . ($i + 3);
        }
        $lines[] = "typeB: -" . ($minBSubjects * 5) . " scale + " . implode(' + ', $terms) . " >= 0 ;\n";
    }

    foreach ($ASubjects as $i) {
        $lines[] = "assign" . ($i + 3) . ": m" . ($i + 3) . " = scale;";
    }

    $zVars = [];
    foreach ($subjects as $i => $s) {
        if (in_array($i, $ASubjects)) continue;
        $m = "m" . ($i + 3);
        $z = "z" . ($i + 3);
        $lines[] = "$m <= 10 $z ;";
        $lines[] = "$m - scale - 10 $z >= -10 ;";
        $lines[] = "$m - scale + 10 $z <= 10 ;";
        $zVars[] = $z;
    }

    $lines[] = "bin " . implode(', ', $zVars) . " ;\n";

    return implode("\n", $lines);
}

//Run worst case LP
file_put_contents('model_min.lp', generateLPMin($subjects, $BSubjects, $ASubjects, 240, 242));
exec("lp_solve model_min.lp", $minOutput, $minStatus);

$worstECTS = 0;
$worstNumerator = 0;

if ($minStatus === 0) {
    $minResult = [];
    $parsing = false;

    foreach ($minOutput as $line) {
        if (!$parsing && str_contains($line, 'Actual values of the variables:')) {
            $parsing = true;
            continue;
        }
        if ($parsing) {
            $line = trim($line);
            if ($line === '') break;
            if (preg_match('/^m(\d+)\s+([\d.]+)/', $line, $m)) {
                $minResult[(int)$m[1]] = (float)$m[2];
            }
        }
    }

    foreach ($subjects as $i => $s) {
        $mVar = $i + 3;
        if (!empty($minResult[$mVar])) {
            $worstECTS += $s['ects'];
            $worstNumerator += $s['mark'] * $s['ects'];
        }
    }

    $worstAvg = $worstNumerator / $worstECTS;
    $gain = $average - $worstAvg;
    echo "<div class='worst-case'>";
    echo "<div><strong>Χειρότερος δυνατός συνδυασμός</strong></div>";
    echo "<div><strong>Συνολικά ECTS:</strong> $worstECTS</div>";
    echo "<div><strong>Βαθμός:</strong> " . round($worstAvg, 2) . "</div>";
    echo "<div class='gain'>Ο καλύτερος δυνατός συνδυασμός αύξησε τον τελικό βαθμό κατά <strong>" . round($gain, 2) . "</strong>.</div>";
    echo "</div>";

    } 
    else {
    echo "<div class='worst-case'><strong>Ο υπολογισμός του χειρότερου συνδυασμού δεν ήταν δυνατός/ Δεν υπάρχει χειρότερος συνδυασμός.</strong></div>";
}


$report = ob_get_clean();
?>


<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Analysis Result</title>
    <style>
    html {
      height: 100%;
    }
    body {
      margin: 0;
      padding: 0;
      min-height: 100%;
      font-family: Arial, sans-serif;
      background: linear-gradient(135deg, #66ccff, #66ff55);
      background-attachment: fixed;
      background-repeat: no-repeat;
      background-size: cover;
    }

    .page-container {
        max-width: 960px;
        margin: 0 auto;
        background-color: white;
        padding: 2em;
        min-height: 100vh;
        box-sizing: border-box;
    }

    h1, h2 {
        text-align: center;
        margin-top: 2em;
    }

    .subject-line {
        display: flex;
        justify-content: space-between;
        max-width: 800px;
        margin: 0.5em auto;
        padding: 0.5em;
        border-bottom: 1px solid #ddd;
    }

    .subject-line .left {
        text-align: left;
    }

    .subject-line .right {
        text-align: right;
        font-weight: bold;
    }

    .excluded {
        color: gray;
        text-align: center;
    }

    .summary {
        font-weight: bold;
        margin-top: 1em;
        text-align: center;
    }

    .worst-case {
    margin-top: 3em;
    padding-top: 1em;
    text-align: left;
    max-width: 800px;
    margin-left: auto;
    margin-right: auto;
}

.gain {
    margin-top: 1em;
    font-weight: bold;
}

    a {
        display: block;
        margin: 2em auto 0 auto;
        text-align: center;
        text-decoration: none;
        font-weight: bold;
    }

    a:hover {
        text-decoration: underline;
    }
    
    .divider {
    border: none;
    border-top: 2px dashed powderblue;
    margin: 3em auto 2em auto;
    width: 80%;
}

</style>


</head>
<body>
    <div class="page-container">
        <h1>Ανάλυση Μαθημάτων</h1>
        <?= $report ?>
        <br><br>
        <a href="index.html">Πίσω στην αρχική σελίδα</a>
    </div>
</body>

</html>

