<?php

class ExcelImportExport {
  
    private $db;

    public function __construct() {
        // Connect to the database using PDO
        $this->db = new PDO('mysql:host=localhost;dbname=database_name', 'username', 'password');
    }
    
    // Method to export data to excel
    public function exportToExcel() {
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
    }

    // Method to import data from excel
    public function importFromExcel($file) {
        // Use PHPExcel library to read the data from the uploaded Excel file
        $excel = PHPExcel_IOFactory::load($file);
        $data = $excel->getActiveSheet()->toArray();

        // Prepare the query to insert the data into the database
        $query = "INSERT INTO table_name (column1, column2, column3) VALUES (:column1, :column2, :column3)";
        $stmt = $this->db->prepare($query);

        // Iterate over the data and execute the query for each row
        foreach ($data as $row) {
            $stmt->execute([
                ':column1' => $row[0],
                ':column2' => $row[1],
                ':column3' => $row[2],
            ]);
        }
    }
    
    // Method to check if a file is excel
    public function isExcel($file) {
        $allowedFileTypes = array('application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        if(in_array($file['type'], $allowedFileTypes)){
            return true;
        }
        return false;
    }
    
}
