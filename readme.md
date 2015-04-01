
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

>in this example `$result` could be a number of things. Some excel libraries such as PHPExcel use their own object to represent a workbook. In that case, a `PHPExcel` object would get returned. Other libraries might return a file pointer, or simply a string that is the file path. Make sure to understand how your particular adapter works.


### Make your own adapter

The majority of your boilerplate code will probably go in the `adapter`. All adapters must implement the `ExcelOutput\ExcelAdapterInterface`. This interface requires the following methods:

* `getFormats()`: should return an array of all supported formats
* `exportToFormat(\ExcelOutput\ExcelWorkbok $workbook, $format, $write = true)`: actually performs the export. Most of the time this will trigger a write to the appropriate file path. However, setting `$write` to false will bypass the write and perform whatever else the given adapter does.
* `supportsFormat($format)`: Simply returns a boolean that indicates if the given format is supported

### Make a Manager

The `manger` is the class applications interact with. You can add any helper methods or extra boilerplate code into the manager. The manager must implement the `ExcelOutput\ExcelManagerInterface` which brings the following restrictions:

* `__construct(ExcelOutput\ExcelAdaperInterface $adapter, ExcelOutput\ExcelFormaterInterface $formatter = null)`: this ensures that the manager will always have access to an adapter and that a `formatter` can be optionally supplied.
* `newWorkbook(array $sheets, Parambag $meta = null)`: creates a new `ExcelOutput\Workbook` object and sets the sheets and any provided meta-data.
* `newSheet(array $data, Parambag $meta = null)`: creates a new `ExcelOutput\SpreadSheet` object with the given data and meta-data
* `formatSheet(ExcelOutput\SpreadSheet $sheet)`: applies format rules found in the given `$sheet`
* `exportToFormat(ExcelOutput\ExcelWorkbook, $format, $write = true)`: actually performs the export. Most of the time this will trigger a write to the appropriate file path. However, setting `$write` to false will bypass the write and perform whatever else the given adapter does.
* `getFormats()`: returns an array of all supported formats.
* `addSheetToWorkbook(ExcelOutput\SpreadSheet $sheet, ExcelOutput\ExcelWorkbook $workbook)`: handles adding a sheet to a work book.
* `supportsFormat($format)`: Simply returns a boolean that indicates if the given format is supported

### Make a formatter (optional)

Sometimes the code required to add format rules to spreadsheets and workbooks can live in the `adapter` or the `manager`. In other cases, it makes more sense to create a `formatter` to handle all of this logic. This will most likely be dictated by how the underpinning library works. Using a `formatter` class to handle applying format rules is entirely up to you.

Should you want to build one, it must use the `ExcelOutput\Formatters\ExcelFormatterInterface` which guarantees the following behavior:

* `applyRules(ExcelOutput\SpreadSheet $sheet)`: applies the rules that apply to the given $sheet

> If you need to inject an object from another library, like a PHPExcel,you can use the formatter's constructor or make a setter for it.

## PHPExcel Integration

Out of the box, you can start working with `PHPExcel` right away. The included `ExcelOutput\PHPExcel` folder has an adapter, formatter, and an abstract manager you can use. Some things to note when using the PHPExcel file:

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
The following objects can handle the given constants as keys and expect a the given input value for the meta-data `Parambag`:

**ExcelWorkbook**
* `ExcelOutput\PHPExcelAdapter::Creator` (string)
* `ExcelOutput\PHPExcelAdapter::LastModifiedBy` (string)
* `ExcelOutput\PHPExcelAdapter::Title` (string)
* `ExcelOutput\PHPExcelAdapter::Subject` (string)
* `ExcelOutput\PHPExcelAdapter::Description` (string)
* `ExcelOutput\PHPExcelAdapter::Keywords` (array<string>)
* `ExcelOutput\PHPExcelAdapter::Category` (string)

**SpreadSheet**
* `ExcelOutput\PHPExcelAdapter::Title` (string)

**Format Rules**
All constants in the following classes are supported:

* `PHPExcel_Style_Alignment`
* `PHPExcel_Style_Border`
* `PHPExcel_Style_Borders`
* `PHPExcel_Style_Color`
* `PHPExcel_Style_Fill`
* `PHPExcel_Style_Font`
* `PHPExcel_Style_NumberFormat`

Additionally the following can be used to set width and height:
*  `ExcelOutput\PHPExcelAdapter::COLUMN_WIDTH` => array(column<string> => width<int>)
*  `ExcelOutput\PHPExcelAdapter::RowHeight` => array(row<int> => height<int>)
*  `ExcelOutput\PHPExcelAdapter::DEFAULT_COLUMN_WIDTH` => width <int>
*  `ExcelOutput\PHPExcelAdapter::DEFAULT_ROW_HEIGHT` => height<int>

### PHPExcel Supported File Formats
* .xlsx: `ExcelOutput\PHPExcelAdapter::XLSX`
* .xls: `ExcelOutput\PHPExcelAdapter::XLS`
* .csv: `Exceloutput\PHPExcelAdapter::CSV`

# to-do's
* add conditional formating support for PHPExcel. needs its own data object in ExcelOutput
* add support for PHPExcel calculation engine
* add support for formulas