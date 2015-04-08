
#ExcelOutput

This library makes use of the [adapter pattern](http://en.wikipedia.org/wiki/Adapter_pattern) to provide applications with a standard interface when it comes to pushing data into a file format supported by Excel. You can still use your library of choice to actually create the Excel files, like [PHPExcel](https://github.com/PHPOffice/PHPExcel). The `ExcelOutput` interfaces reduce the footprint of boilerplate code required to interact with the excel library of your choice. It will also become easier to switch to a new library, or even to use multiple libraries side by side if you choose. 

## Note about structure

`ExcelOutput` uses three data objects to keep track of everything:

1. `ExcelOutput\ExcelWorkbook`: used to keep track of all information that will go into an excel workbook.
2. `ExcelOutput\SpreadSheet`: used to keep track of all data and meta-data that will go into a spread sheet. Also contains an array of *FormatRules*.
3. `ExcelOutput\FormatRule`: keeps track of a column range and format rule to apply. The format rule is simply an array that will get passed through. PHPExcel's documentation details how to create a $style array. Different libraries might require something else Excel accepts the following ranges
    * `A1`: single cell
    * `A1:A8`: range of cells
    * `A:C`: range of columns
*Please note that it is best to only have one FormatRule object per range.*

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

>in this example `$result` could be a number of things. Some excel libraries such as PHPExcel use their own object to represent a workbook. In that case, a `PHPExcel` object would get returned. Other libraries might return a file pointer, or simply a string that is the file path. Make sure to understand how your particular adapter works.


### Make your own adapter

The majority of your boilerplate code will probably go in the `adapter`. All adapters must implement the `ExcelOutput\ExcelAdapterInterface`. This interface requires the following methods:

* `getFormats()`: should return an array of all supported formats
* `exportToFormat(\ExcelOutput\ExcelWorkbok $workbook, $format, $write = true)`: actually performs the export. Most of the time this will trigger a write to the appropriate file path. However, setting `$write` to false will bypass the write and perform whatever else the given adapter does.
* `supportsFormat($format)`: Simply returns a boolean that indicates if the given format is supported

### Make a Manager

The `manger` is the class applications interact with. You can add any helper methods or extra boilerplate code into the manager. The manager must implement the `ExcelOutput\ExcelManagerInterface` which brings the following restrictions:

* `__construct(ExcelOutput\ExcelAdaperInterface $adapter, ExcelOutput\ExcelFormaterInterface $formatter = null)`: this ensures that the manager will always have access to an adapter and that a `formatter` can be optionally supplied.
* `newWorkbook(array $sheets, array $meta = null)`: creates a new `ExcelOutput\Workbook` object and sets the sheets and any provided meta-data.
* `newSheet(array $data, array $meta = null)`: creates a new `ExcelOutput\SpreadSheet` object with the given data and meta-data
* `formatSheet(ExcelOutput\SpreadSheet $sheet)`: applies format rules found in the given `$sheet`
* `exportToFormat(ExcelOutput\ExcelWorkbook, $format, $write = true)`: actually performs the export. Most of the time this will trigger a write to the appropriate file path. However, setting `$write` to false will bypass the write and perform whatever else the given adapter does.
* `getFormats()`: returns an array of all supported formats.
* `addSheetToWorkbook(ExcelOutput\SpreadSheet $sheet, ExcelOutput\ExcelWorkbook $workbook)`: handles adding a sheet to a work book.
* `supportsFormat($format)`: Simply returns a boolean that indicates if the given format is supported

### Make a formatter (optional)

Sometimes the code required to add format rules to spreadsheets and workbooks can live in the `adapter` or the `manager`. In other cases, it makes more sense to create a `formatter` to handle all of this logic. This will most likely be dictated by how the underpinning library works. Using a `formatter` class to handle applying format rules is entirely up to you.

Should you want to build one, it must use the `ExcelOutput\Formatters\ExcelFormatterInterface` which guarantees the following behavior:

* `applyRules(ExcelOutput\SpreadSheet $sheet, $target = null`: applies the rules that apply to the given $sheet

> If you need to inject an object from another library, like a PHPExcel,you can pass that object in as the $target.

## PHPExcel Integration

Out of the box, you can start working with `PHPExcel` right away. The included `ExcelOutput\PHPExcel` folder has an adapter, formatter, and an abstract manager you can use. Some things to note when using the PHPExcel file:

### PHPExcel Formatter / Adapter

Because of the way PHPExcel works, an instance of `ExcelOutput\PHPExcel\PHPExcelFormatter` is required by the adapter. PHPExcel itself requires the proprietary PHPExcel object to get instantiated before applying format rules. This means that the `ExcelOutput\PHPExcel\PHPExcelAdapter` **requires** the formatter. Additionally, the `ExcelOutput\PHPExcel\PHPExcelManager` also **requires** a formatter to fulfill all of its obligations Simply instantiate the formatter first and pass it to both constructors:

```php
$formatter = new PHPExcelFormatter();
$adapter = new PHPExcelAdapter($formatter);
// you can use a different formatter instance, or another object that implements the correct interface if you choose. 
$manager = new PHPExcelManager($adapter, $formatter);
```

### PHPExcel Format Rules

PHPExcel uses different methods to set different style rules. As a result, the `Excloutput\PHPExcel\PHPExcelFormatter` expects each format rule to have the correct `$type` attribute. To make things easier, a `ExcelOutput\PHPExcel\FormatRuleFactory` was included that allows you to create correctly configured formatting rules. Each method accepts a range and style argument. In all cases, the range must be a valid Excel column / cell range. The style argument must be applicable to the context of the called method:

* `createStyleRule($range, $style)`: $style is a valid $styleArray as defined in PHPExcel's Documentation
* `createNumFormatRule($range, $style)`: $style is a valid number format from the *PHPExcel_STYLE_NUMBERFORMAT* class.

### Caching strategies

PHPExcel supports the use of several different caching strategies. You can read about the specific strategies themselves in the PHPExcel documentation. List of available caching strategies:

* Memory: `PHPExcel_CachedObjectStorageFactory::cache_in_memory`   (**default**)
* Memory Gzip: `PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip`
* Memory Serialized: `PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized`
* Memory IgBinary: `PHPExcel_CachedObjectStorageFactory::cache_igbinary`
* DiscISAM: `PHPExcel_CachedObjectStorageFactory::cache_to_discISAM`
* APC: `PHPExcel_CachedObjectStorageFactory::cache_to_apc`
* Memcache: `PHPExcel_CachedObjectStorageFactory::cache_to_memcache`
* PHPTemp: `PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp`
* Wincache: `PHPExcel_CachedObjectStorageFactory::cache_to_wincache`
* sqlite: `PHPExcel_CachedObjectStorageFactory::cache_to_sqlite`
* sqlite3: `PHPExcel_CachedObjectStorageFactory::cache_to_sqlite3`

To set a caching method:

```php
$adapter = new PHPExcelAdapter();
$manager = new PHPExcelManager($adapter);

$memcacheOptions = array(
    'memcachedServer' => 'localhost',
    'memcachePort'    => 11211,
    'cacheTime'       => 600 
);
$cacheStrategy = array(
    PHPExcel_CachedObjectStorageFactory::cache_to_memcache => $memcacheOptions

$manger->setCacheStrategy($cacheStrategy)

```
>If the strategy you choose requires extra options, like Memcache, just include the options array as the value and the strategy as the key. If you do not need extra options, just include the strategy as the value

### Valid Meta-Data
The following objects can handle the given constants as keys and expect a the given input value for the meta-data array:

**ExcelWorkbook**
* `ExcelOutput\PHPExcelAdapter::CREATOR` (string)
* `ExcelOutput\PHPExcelAdapter::LAST_MODIFIED_BY` (string)
* `ExcelOutput\PHPExcelAdapter::TITLE` (string)
* `ExcelOutput\PHPExcelAdapter::SUBJECT` (string)
* `ExcelOutput\PHPExcelAdapter::DESCRIPTION` (string)
* `ExcelOutput\PHPExcelAdapter::KEYWORDS` (array<string>)
* `ExcelOutput\PHPExcelAdapter::CATEGORIES` (array<string>)

**SpreadSheet**
* `ExcelOutput\PHPExcelAdapter::Title` (string)

### PHPExcel Supported File Formats
* .xlsx: `ExcelOutput\PHPExcelAdapter::XLSX`
* .xls: `ExcelOutput\PHPExcelAdapter::XLS`
* .csv: `Exceloutput\PHPExcelAdapter::CSV`

# to-do's
* add conditional formating support for PHPExcel
* add support for additional formating stuff not in a PHPExcel Style Array.
* add support for PHPExcel calculation engine
* add support for formulas