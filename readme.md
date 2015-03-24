# Excel Output

This library will use the (adapter)[http://en.wikipedia.org/wiki/Adapter_pattern] pattern to create an easy way to export data to excel. Out of the box, you will only have access to one adapter which uses (PHPExcel)[https://github.com/PHPOffice/PHPExcel] under the hood. If you prefer some other library over PHPExcel, you will just need to build your own adapter by implementing the correct interfaces. 

## Data Objects

Two data objects exist to act as a common way to describe workbooks and the spreadsheets they contain. `ExcelWorkbook` and `ExcelSheet` contain all of the information that describes their excel counterparts, such as data for tables, formatting rules, as well as other bits of meta-data. You can interact with these objects directly, or you can use a simple array structure and the appropriate factory to create these objects.

### ExcelWorkbook Options

ExcelWorkbooks can have the following options:

* Title, string
* Creator, string
* Last Modified, DateTime
* Subject, string
* Description, string
* Keywords, array<string>
* Category, string
* Company, string

The workbook class contains constants for each of these values. The constants should get used as array keys in an `$options` array used to create an `ExcelWorkbook`. All of them are optional. You can use any of them, none of them, or all of them. Example:

```
$workBookOptions = array(
    ExcelWorkbook::META_TITLE => 'Fools I Pity',
    ExcelWorkbook::META_CREATOR => 'Mr. T',
    ExcelWorkbook::META_LAST_MODIFIED => new \DateTime(),
    ExcelWorkbook::META_SUBJECT => 'Suckas',
    ExcelWorkbook::META_DESCRIPTION => 'Who I pity, how much and why',
    ExcelWorkbook::META_KEYWORDS => array('fools', 'pity'),
    ExcelWorkbook::META_CATEGORY => 'Personal',
    ExcelWorkbook::META_COMPANY => 'A-Team'
);  
```

### Sheet Options

Excel Sheets have two types of options, meta-data and formatting rules. The meta-data simply describes the sheet itself. Formating rules get used to control the actual display of the sheet. These rules get applied during the `export` process.

**Meta Data Options**

* Display Name, string (display value for spreadsheet in excel tab)
* Alias, string (optional, defaults to display name. Used to programatically refer to the spreadsheet).

**Format Options**
Each format option contains the range of cells that the rule applies to, and the actual value of the rule itself. Ranges formats supported include just cells: A1:A7, entire columns: A:D, or even a single cell: A1.

* Number Format, array<ragnge => format> (see valid number formats)
* Background Color, array<range => hexCode>
* Column Width, array<range => int>

**Example:**

```
$sheetsOption = array(
    ExcelSheet::META = array(
        ExcelSheet::DISPLAY_NAME => '',
        ExcelSheet::ALIAS => ''
        ),
    ExcelSheet:FORMAT = array(
        ExcelSheet::NUMBER_FORMAT => array('A:B' => ExcelSheet:FORMAT_CURRENCY),
        ExcelSheet::BACKGROUND_COLOR => array('A1' => 'dce8f2'),
        ExcelSheet::COLUMN_WIDTH => array(A:C => 20)
    )
);
```

## Adapters

All boilerplate code used to interact with a particular library, such as PHPExcel, will exist in an adapter. Adapters **must** implement the `Mbright\ExcelOutput\AdapterInterface`. This will allow applications to use whatever excel wrapper they choose through a common interface. Most applications will probably only use one adapter, though you could use as many as you want. Out of the box you have access to an adapter for (PHPExcel)[https://github.com/PHPOffice/PHPExcel].

## Manager

Your application should use the manager as the single point of contact. Each manager has its own adapter, and knows how to use the provided factories to create `ExcelWorkbook` and `ExcelSheet` objects. You can either manipulate those objects by hand, or give the manager a predefined array structure and have it do all of the hard work for you. If you decide to use a custom adapter, the default manager will work just fine. It merely relies on the `ExcelOutput\AdapterInterface`. Just give the manager whichever adapter you want to use.

## Usage

Create an `ExcelOutput\ExcelWorkbook` object:
```
$adapter = new PHPExcelAdapter();
$manager = new ExcelManager($adapter);

$workBookOptions = array(
    ExcelWorkbook::META_TITLE => 'Fools I Pity',
    ExcelWorkbook::META_CREATOR => 'Mr. T',
    ExcelWorkbook::META_LAST_MODIFIED => new \DateTime(),
    ExcelWorkbook::META_SUBJECT => 'Suckas',
    ExcelWorkbook::META_DESCRIPTION => 'Who I pity, how much and why',
    ExcelWorkbook::META_KEYWORDS => array('fools', 'pity'),
    ExcelWorkbook::META_CATEGORY => 'Personal',
    ExcelWorkbook::META_COMPANY => 'A-Team'
); 

$workBook = $manager->newWorkBook($options);
```

Create a `ExcelOutupt\ExcelSheet` object:

```
$adapter = new PHPExcelAdapter();
$manager = new ExcelManager($adapter);

$sheetsOption = array(
    ExcelSheet::META = array(
        ExcelSheet::DISPLAY_NAME => '',
        ExcelSheet::ALIAS => ''
        ),
    ExcelSheet:FORMAT = array(
        ExcelSheet::NUMBER_FORMAT => array('A:B' => ExcelSheet:FORMAT_CURRENCY),
        ExcelSheet::BACKGROUND_COLOR => array('A1' => 'dce8f2'),
        ExcelSheet::COLUMN_WIDTH => array(A:C => 20)
    )
);

$excelSheet = $manager->newSheet($options);
```

Add data to a sheet object:

```
$data = array(1,2,3,4,5);

$excelSheet->setData($data);
```

Add a sheet to a WorkBook:

```
$workBook->addSheet($excelSheet)

//multiple sheets
$workBook->addSheets(array($excelSheet, $excelSheet2));
```

Export a WorkBook:

```
$destination = 'path/to/my/spreadsheet.xlsx';
$format = 'Excel2007';

$manager->write($workBook, $destination, $format);
```

Create a workbook and its sheets in one step:

```
$adapter = new PHPExcelAdapter();
$manager = new ExcelManager($adapter);

$workBookOptions = array(
    ExcelWorkbook::META_TITLE => 'Fools I Pity',
    ExcelWorkbook::META_CREATOR => 'Mr. T',
    ExcelWorkbook::META_LAST_MODIFIED => new \DateTime(),
    ExcelWorkbook::META_SUBJECT => 'Suckas',
    ExcelWorkbook::META_DESCRIPTION => 'Who I pity, how much and why',
    ExcelWorkbook::META_KEYWORDS => array('fools', 'pity'),
    ExcelWorkbook::META_CATEGORY => 'Personal',
    ExcelWorkbook::META_COMPANY => 'A-Team'
); 

$sheet1Option = array(
    ExcelSheet::META = array(
        ExcelSheet::DISPLAY_NAME => 'Sheet 1',
        ExcelSheet::ALIAS => 'first'
        ),
    ExcelSheet:FORMAT = array(
        ExcelSheet::NUMBER_FORMAT => array('A:B' => ExcelSheet:FORMAT_CURRENCY),
        ExcelSheet::BACKGROUND_COLOR => array('A1' => 'dce8f2'),
        ExcelSheet::COLUMN_WIDTH => array(A:C => 20)
    )
);

$sheet2Option = array(
    ExcelSheet::META = array(
        ExcelSheet::DISPLAY_NAME => 'Sheet 2',
        ExcelSheet::ALIAS => 'second'
        ),
    ExcelSheet:FORMAT = array(
        ExcelSheet::NUMBER_FORMAT => array('A:B' => ExcelSheet:FORMAT_CURRENCY),
        ExcelSheet::BACKGROUND_COLOR => array('A1' => 'dce8f2'),
        ExcelSheet::COLUMN_WIDTH => array(A:C => 20)
    )
);
$sheetsOptions = array($sheet1Option, $shet2Optione);

$data = array(1,2,3,4,5);
$data2 = array(6,7,8,9,10);
$dataForSheets = array($data, $data2);

$workBook = $manager->newWorkBookWithData($data, $workBookOptions, $sheetOptions);
```

## PHPExcel Adapter

PHPExcel functions by creating a representation of your spreadsheet in memory. It offers several (caching strategies)[https://github.com/PHPOffice/PHPExcel/blob/develop/Documentation/markdown/Overview/04-Configuration-Settings.md] that you can use to better fit your needs. 

How to set the cache strategy:

```
$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory;
$adapter = new PHPExcelAdapter($cacheMeethod);
```


