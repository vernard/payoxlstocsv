<?php

namespace PayoXlsToCsv;

class PayoXlsToCsv {
    protected $batchSize = 500;

    public function __construct($output_dir) {
        $this->output_dir = $output_dir;
    }

    public function convert($file) {
        $processed = 0;
        do{
            unset($verifiedData);
            // Keep fetching the records, row per row. Output them for every 500 records to stop memory overload
            $verifiedData = $this->_fetchRecords($processed, $file);
            // Write records on the CSV
            $this->_appendRecords($verifiedData, $file);
            $processed += count($verifiedData);
        } while (count($verifiedData) >= $this->batchSize);
    }

    protected function _fetchRecords(&$processed, $file) {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);
        $chunkFilter = new \PayoXlsToCsv\ChunkReadFilter();
        $reader->setReadFilter($chunkFilter);


        $startRow = $processed + 11;
        $endRow = $processed + $this->batchSize + 10;
        $chunkFilter->setRows($startRow,$this->batchSize);
        $spreadsheet = $reader->load($file);
        $worksheet = $spreadsheet->getActiveSheet();

        $dataArray = $worksheet
            ->rangeToArray("B{$startRow}:U{$endRow}",null,
                false,false,false);

        $dataArray = $this->_removeMergedColumn($dataArray);

        return $this->_cutEndData($dataArray);
    }

    protected function _cutEndData(array $arr) {
        for($x = count($arr)-1; $x>0; $x--) {
            if($arr[$x][0] == "Total") {
                unset($arr[$x]);
                break;
            }
            elseif(! is_null($arr[$x][0])) {
                break;
            }
            elseif(is_null($arr[$x][0])) {
                unset($arr[$x]);
            }
        }
        return $arr;
    }

    protected function _appendRecords($verifiedData, $xlsFile) {
        // Create a equivalent CSV file if not exists
        $csvFile = str_replace(".xlsx", ".csv", basename($xlsFile));
        $csvFile = $this->output_dir . $csvFile;
        if(! file_exists($csvFile)) {
            file_put_contents($csvFile, "Shipbill ID,Order ID,Created At,Pickup Date,Customer Name,Order Amount,Payo Expected Shipping Fee,Payo Actual Shipping Fee,Payo Service Status,Payo Courier Status,Payo Pending Reason,Mobile Number,Address,Tracking Number,Payo Notes,Payo Courier Notes,Courier,Payo Last Status Update,Payo Shipbill Url\n");
        }

        // Load the equivalent CSV file and append
        $handle = fopen($csvFile, "a");
        foreach($verifiedData as $row) {
            fputcsv($handle, array_values($row));
        }
        fclose($handle);
    }

    protected function _removeMergedColumn(array $dataArray) {
        foreach($dataArray as $key => $value) {
            unset($dataArray[$key][1]);
        }

        return $dataArray;
    }
}