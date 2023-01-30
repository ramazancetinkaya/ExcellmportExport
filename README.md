# ExcellmportExport
Excel Import/Export

To use the ExcelImportExport class, you need to first create an instance of the class, and then call its methods as needed.

## Here is an example of how you could use the class to export data to an Excel file:
```php
require_once 'path/to/ExcelImportExport.php';

$excel = new ExcelImportExport();
$excel->exportToExcel();
```

## And an example of how you could use the class to import data from an Excel file:
```php
require_once 'path/to/ExcelImportExport.php';

$excel = new ExcelImportExport();

if(isset($_FILES['file'])){
    $file = $_FILES['file'];
    if($excel->isExcel($file)){
        $excel->importFromExcel($file['tmp_name']);
    }
}
```

You can also use the isExcel() function to check if the file uploaded is an excel sheet before processing it.

It's important to note that you need to include the PHPExcel library in your project in order to use the export and import functionality.

Also, you need to make sure that the database connection details in the class constructor match the details of your database.

### Authors

**Ramazan Çetinkaya**

* [github/ramazancetinkaya](https://github.com/ramazancetinkaya)

### License

Copyright © 2023, [Ramazan Çetinkaya](https://github.com/ramazancetinkaya).
