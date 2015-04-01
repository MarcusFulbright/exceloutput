
#ExcelOutput

This library makes use of the [adapter pattern](http://en.wikipedia.org/wiki/Adapter_pattern) to provide applications with a standard interface when it comes to pushing data into a file format supported by Excel. You can still use your library of choice to actually create the Excel files, like [PHPExcel](https://github.com/PHPOffice/PHPExcel). The `ExcelOutput` interfaces reduce the footprint of boilerplate code required to interact with the excel library of your choice. It will also become easier to switch to a new library, or even to use multiple libraries side by side if you choose. 

## Note about structure

`ExcelOutput` uses three data objects to keep track of everything:

1. `ExcelOutput\ExcelWorkbook`: used to keep track of all information that will go into an excel workbook.
2. `ExcelOutput\SpreadSheet`: used to keep track of all data and meta-data that will go into a spread sheet. Also contains an array of *FormatRules*.
3. `ExcelOutput\FormatRule`: keeps track of a column range and format rule to apply. Excel accepts the following ranges
    * `A1`: single cell
    * `A1:A8`: range of cells
    * `A:C`: range of columns 

## Usage

Generally speaking your workflow will probably look like this:

* Instantiate your `adapter`
* Instantiate a `manager` and pass it the `adapter`
* Create the needed `ExcelOutput\Workbook` objects
* Create the needed `ExcelOutput\Spreadsheet` objects
* Attach sheets to the appropriate workbook(s)
* Export the workbook to the desired format

Example:

```php
$adapter = new MyAdapter();
$manager = new MyManager($adapter);

$workbook = $manager->createWorkbook('My Workbook');
$sheet = $manager->createSheet('Sheet 1',array(1,2,3,4,5));

$manager->addSheetToWorkbook($sheet, $workbook);
$result = $manager->exportWorkbook($workbook, 'Excel2007');
```

>in this example `$result` could be a number of things. Some excel libraries such as PHPExcel use their own object to represent a workbook. In that case, a `PHPExcel` objct would get returned. Other libraries might return a file pointer, or simply a string that is the file path. Make sure to understand how your particular adapter works.


### Make your own adapter

The majority of your boilerplate code will probably go in the `adapter`. All adapers must implement the `ExcelOutput\Adapters\ExcelAdapterInterface`. This interface requires the following methods:

* `getFormats()`: should return an array of all supported formats
* `exportToFormat(\ExcelOutput\ExcelWorkbok $workbook, $format, $write = true)`: actually performs the export. Most of the time this will trigger a write to the appropriate file path. However, setting `$write` to false will bypass the write and perform whatever else the given adapter does.
* `supportsFormat($format)`: Simply returns a boolean that indicates if the given format is supported

### Make a Manager

The `manger` is the class applications interact with. You can add any helper methods or extra boilerplate code into the manager. The manager must implement the `ExcelOutput\Manager\ExcelManagerInterface` which brings the following restrictions:

* `__construct(ExcelOutput\Adapters\ExcelAdaperInterface $adapter, ExcelOutput\Formattters\ExcelFormaterInterface $formatter = null)`: this ensures that the manager will always have access to an adapter and that a `formatter` can be optionally supplied.
* `newWorkbook($name, array $properties = null)`: creates a new `ExcelOutput\Workbook` object and sets the provided properties.
* `newSheet(array $data, array $prpoerties = array)`: creates a new `ExcelOutput\ExcelSheet` $object with the given data and sets the provided properties.
* `formatSheet(ExcelOutput\ExcelSheet $sheet)`: applies format rules found in the given `$sheet`
* `exportToFormat(ExcelOutput\ExcelWorkbook, $format, $write = true)`: actually performs the export. Most of the time this will trigger a write to the appropriate file path. However, setting `$write` to false will bypass the write and perform whatever else the given adapter does.
* `getFormats()`: returns an array of all supported formats.
* `addSheetToWorkbook(ExcelOutput\ExcelSheet $sheet, ExcelOutput\ExcelWorkbook $workbook)`: handles adding a sheet to a work book.
* `supportsFormat($format)`: Simply returns a boolean that indicates if the given format is supported

### Make a formatter (optional)

Sometimes the code required to add format rules to spreadsheets and workbooks can live in the `adapter` or the `manager`. In other cases, it makes more sense to create a `formatter` to handle all of this logic. This will most likely be dictated by how the underpinning library works. Using a `formater` class to handle applying format rules is entirely up to you.

Should you want to build one, it must use the `ExcelOutput\Formatters\ExcelFormatterInterface` which guarntees the following behavior:

* `applyRules(ExcelOutput\SpreadSheet $sheet)`: applies the rules that apply to the given $sheet

> If you need to inject an object from another library, like a PHPExcel,you can use the formatter's constructor or make a setter for it.
