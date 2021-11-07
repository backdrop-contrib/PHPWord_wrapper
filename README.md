# PHPWord Wrapper

This module is simply a wrapper around the PHPWord and PHPOffice-Common php libraries. It allows Backdrop programmers to generate MS Word and other files through coding.

NOTE: This module uses library PHPWord-0.18.2

NOTE 2: I have added a custom Autoloader.php file to the library, located under `libraries/PHPWord/src/Phpword/Autoloader.php`

## Installation

  - Install this module using the official Backdrop CMS instructions at
https://backdropcms.org/guide/modules

## Example Use

In your custom module, you may use the following sample code to generate sample files.

```php
  $phpWord = new \PhpOffice\PhpWord\PhpWord();
  \PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true); // This automatically escapes ampersands etc.

  /* Note: any element you append to a document must reside inside of a Section. */

  // Adding an empty Section to the document...
  $section = $phpWord->addSection();
  // Adding Text element to the Section having font styled by default...
  $section->addText(
      '"Learn from yesterday, live for today, hope for tomorrow. '
          . 'The important thing is not to stop questioning." '
          . '(Albert Einstein)'
  );

  /*
  * Note: it's possible to customize font style of the Text element you add in three ways:
  * - inline;
  * - using named font style (new font style object will be implicitly created);
  * - using explicitly created font style object.
  */

  // Adding Text element with font customized inline...
  $section->addText(
      '"Great achievement is usually born of great sacrifice, '
          . 'and is never the result of selfishness." '
          . '(Napoleon Hill)',
      array('name' => 'Tahoma', 'size' => 10)
  );

  // Adding Text element with font customized using named font style...
  $fontStyleName = 'oneUserDefinedStyle';
  $phpWord->addFontStyle(
      $fontStyleName,
      array('name' => 'Tahoma', 'size' => 10, 'color' => '1B2232', 'bold' => true)
  );
  $section->addText(
      '"The greatest accomplishment is not in never falling, '
          . 'but in rising again after you fall." '
          . '(Vince Lombardi)',
      $fontStyleName
  );

  // Adding Text element with font customized using explicitly created font style object...
  $fontStyle = new \PhpOffice\PhpWord\Style\Font();
  $fontStyle->setBold(true);
  $fontStyle->setName('Tahoma');
  $fontStyle->setSize(13);
  $myTextElement = $section->addText('"Believe you can and you\'re halfway there." (Theodor Roosevelt)');
  $myTextElement->setFontStyle($fontStyle);

  $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007', $download = true);

  // The following code will force the browser to download the Word file
  backdrop_add_http_header("Content-Disposition",  'attachment; filename="myFile.docx"');
  backdrop_add_http_header("Content-type", 'application/octet-stream');
  $objWriter->save("php://output");
  backdrop_exit();
```

Current Maintainers
-------------------

- [argiepiano](https://github.com/argiepiano).

License
-------

This project is GPL v2 software. 
See the LICENSE.txt file in this directory for complete text.

