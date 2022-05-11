# office-to-pdf

Convert any office files to PDF format.

This library use LibreOffice power to convert any office files to PDF format. So, please make sure you have installed LibreOffice on your environment!

## Requirements
- PHP 5.6 or above
- LibreOffice

## Installation

```bash
composer require ngekoding\office-to-pdf
```

## Usage

### Simple usage

```php
<?php

// Include autoloader
require 'vendor/autoload.php';

use Ngekoding\OfficeToPdf\Converter;

$converter = new Converter();
$converter->convert('path/to/file.docx');
```

You can also convert other file type (not only .docx), feel free to convert any office file like `.pptx`, `.xlsx`, etc...

### Setting manually LibreOffice executable path

By default the library will try to find LibreOffice executable based on current operating system.

```php
$converter = new Converter('path/to/libreoffice');
```

### Setting destination folder

By default the PDF result will generated in the same directory as the source file. You can define the destination folder by passing the second parameter.

```php
$converter = new Converter();
$converter->convert('path/to/file.docx', './pdf-outputs');
```

By default the PDF filename will same with the source file, you can define the output filename by setting the destination folder with filename.

```php
// The result filename will be result.pdf in the same directory as the source file
$converter->convert('path/to/file.docx', 'result.pdf');

// Or save to spesific directory
$converter->convert('path/to/file.docx', './pdf-outputs/result.pdf');
```
