<?php

/**
 * ExcelmportExport
 * © 2023 RAMAZAN ÇETİNKAYA, All rights reserved.
 *
 * @author [ramazancetinkaya]
 * @date [24.01.2023]
 *
 * Please note, this class is only a demonstration and should be used with caution, you should test it before using on a production environment.
 */

class ExcelImportExport {
    private $db;

    public function __construct() {
        try {
            // Connect to the database using PDO
            $this->db = new PDO('mysql:host=localhost;dbname=database_name', 'username', 'password');
        } catch (PDOException $e) {
            die('Could not connect to the database: ' . $e->getMessage());
        }
    }

    public function exportToExcel() {
        try {
            // Write the query to select the data you want to export
            $query = "SELECT * FROM table_name";

            // Prepare and execute the query
            $stmt = $this->db->prepare($query);
            $stmt->execute();

            // Fetch the data as an associative array
            $data = $stmt->fetchAll(PDO::FETCH_ASSOC);

            // Use PHPExcel library to create and export the Excel file
            $excel = new PHPExcel();
            $excel->setActiveSheetIndex(0);
            $excel->getActiveSheet()->fromArray($data, null, 'A1');
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="data.xls"');
            header('Cache-Control: max-age=0');
            $writer = PHPExcel_IOFactory::createWriter($excel, 'Excel5');
            $writer->save('php://output');
        } catch (PDOException $e) {
            die('Could not export data to Excel: ' . $e->getMessage());
        }
    }

    public function importFromExcel($file) {
        try {
            // Use PHPExcel library to read the data from the uploaded Excel file
            $excel = PHPExcel_IOFactory::load($file);
            $data = $excel->getActiveSheet()->toArray();

            // Prepare the query to insert the data into the database
            $query = "INSERT INTO table_name (column1, column2, column3) VALUES (:column1, :column2, :column3)";
            $stmt = $this->db->prepare($query);

            // Iterate over the data and execute the query for each row
            foreach ($data as $row) {
                // Bind the values to prevent SQL injection
                $stmt->bindValue(':column1', $row[0]);
                $stmt->bindValue(':column2', $row[1]);
                $stmt->bindValue(':column3', $row[2]);
                $stmt->execute();
            }
        } catch (PDOException $e) {
            die('Could not import data from Excel: ' . $e->getMessage());
        }
    }

    public function isExcel($file) {
        $allowedFileTypes = array('application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        if(in_array($file['type'], $allowedFileTypes)){
            return true;
        }
        return false;
    }
  
}
