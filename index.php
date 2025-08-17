<?php

use PhpOffice\PhpSpreadsheet\Calculation\TextData\Format;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpWord\TemplateProcessor;

require __DIR__ . "/vendor/autoload.php";

$options = getopt('f:iah', ['file:', 'micros', 'macros', 'help']);

if (isset($options['h']) || isset($options['help'])) {

    printf("
    -f or --file => results excel file
    -i or --micros => micros results
    -a or --macros => macros results
    -h or --help => help\n");
    exit;
}

$reader = new Xlsx;
if (!isset($options['f']) && !isset($options['file'])) {
    print("Excel input file name is required. -f/--file \n");
    exit;
}

$spreadsheet = isset($options['f']) ?
    $reader->load(__DIR__ . "/" . $options['f']) :
    $reader->load(__DIR__ . "/" . $options['file']);
$activesheet = $spreadsheet->getActiveSheet();


$results = [];


/** find batches */
$batches = [];

foreach ($activesheet->getRowIterator(2) as $row) {
    $sampleID = ($activesheet->getCell("C" . $row->getRowIndex()))->getValue();
    $batch = str_contains($sampleID, '/') ? explode('/', $sampleID)[0] : $sampleID;

    if (str_contains(strtolower($batch), 'd') && !str_contains(strtolower($batch), '/')) {
        $batch = explode("/", str_replace('d', '/d', strtolower($batch)))[0];
    }

    if (!in_array($batch, $batches)) {

        $batches[$batch] = [];
    }
}

foreach ($activesheet->getRowIterator(2) as $row) {
    $sampleID = ($activesheet->getCell("C" . $row->getRowIndex()))->getValue();
    $batch = str_contains($sampleID, '/') ? explode('/', $sampleID)[0] : $sampleID;
    if (str_contains(strtolower($batch), 'd') && !str_contains(strtolower($batch), '/')) {
        $batch = explode("/", str_replace('d', '/d', strtolower($batch)))[0];
    }

    if (str_contains(strtolower($sampleID), 'd') && !str_contains(strtolower($sampleID), '/')) {
        $sampleID = str_replace('d', '/d', strtolower($sampleID));
    }
    $$sampleID = [
        'Calcium' => $activesheet->getCell('H' . $row->getRowIndex())->getValue(),
        'Potassium' => $activesheet->getCell('K' . $row->getRowIndex())->getValue(),
        'Magnesium' => $activesheet->getCell('O' . $row->getRowIndex())->getValue(),
        'Sodium' => $activesheet->getCell('Q' . $row->getRowIndex())->getValue(),
    ];

    $batches[$batch][$sampleID] = $$sampleID;
}

/** replace high values with diluted values */
$upper_limit = 50;
foreach ($batches as $id => $batch) {
    $sample = '';
    $test = '';

    foreach ($batch as $sample_id => $sample) {
        foreach ($sample as $test => $value) {
            if (!str_contains($value, 'H') && !str_contains($value, '!')) {
                continue;
            }
            $dilutions = [];

            foreach ($batch as $sample_id2 => $sample2) {
                if (!str_contains($sample_id2, $sample_id) || !str_contains(strtolower($sample_id2), 'x')) {
                    continue;
                }
                if (str_contains($sample2[$test], 'H') || str_contains($sample2[$test], '!')) {
                    continue;
                }
                array_push($dilutions, (float)$sample2[$test]);
            }

            if (!empty($dilutions)) {

                $batches[$id][$sample_id][$test] = max($dilutions);
            }
        }
    }
}
/** replace high values with diluted values */

$metals_type = "macros"; // macros or micros

$template_name = $metals_type == "macros" ? __DIR__ . "/worksheet-templates/CATIONS (Macros)   ED03.docx" : __DIR__ . "/worksheet-templates/CATIONS (Micros)   ED03.docx";

foreach ($batches as $batchID => $batch) {

    $template = new TemplateProcessor($template_name);
    $template->setValues([
        "registration_no" => "2025//$batchID",
        "analyst" => "",
        "pg" => "1",
        "pgOf" => "1",
        "pipette_id" => "Gen 036",
        "date" => "",
        "sample_no" => "\${" . "sample_no" . "}"
    ]);

    $template->cloneRow('sample_no', count($batch));

    $key = 1;
    $samples_in_batch = count($batch);

    foreach ($batch as $id => $sample) {
        if (str_contains($id, 'x') || str_contains($id, 'X')) {

            $template->setValues([
                'sample_no#' .  $key => "",
                "cal#" . $key => "",
                "pot#" . $key => "",
                "mag#" . $key => "",
                "sod#" . $key => ""
            ]);
        } else {
            $template->setValues([
                'sample_no#' .  $key => $id,
                "cal#" . $key => $sample['Calcium'],
                "pot#" . $key => $sample['Potassium'],
                "mag#" . $key => $sample['Magnesium'],
                "sod#" . $key => $sample['Sodium']
            ]);
        }


        if (str_contains($id, '/')) {
            $id_split = explode('/', $id);

            foreach ($id_split as $split) {
                if (str_contains(strtolower($split), 'x')) {
                    $template->setValues([
                        "cal_dil#" . $key => "",
                        "pot_dil#" . $key => "",
                        "mag_dil#" . $key => "",
                        "sod_dil#" . $key => ""
                    ]);
                }
            }
        }

        if ($key == $samples_in_batch) {
            $today_worksheets = date('d_m_Y');

            if (!is_dir(__DIR__ . "/worksheets")) {
                mkdir(__DIR__ . "/worksheets");
            }

            if (!is_dir(__DIR__ . "/worksheets/$today_worksheets")) {

                mkdir(__DIR__ . "/worksheets/$today_worksheets");
            }

            foreach ($template->getVariables() as $var) {
                $template->setValue($var, "");
            }
            $template->saveAs(__DIR__ . "/worksheets/$today_worksheets/$batchID.docx");
            break;
        }
        $key++;
    }
}


$file = fopen('Results.json', 'w');
fputs($file, json_encode($batches));
fclose($file);
return true;
