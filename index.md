---
title: "Excel - VBA - Call a web service"
subtitle: "How to"
date: "<!-- concat-md::date 'dd MMMM yyyy, HH:mm' -->"
keywords: []
language: "en"
---
<!-- markdownlint-disable MD025 -->

# Excel - VBA - Call a web service

> How to call a web service in Excel VBA. The example is built using the European Union's VIES CheckVAT web service.

<!-- concat-md::toc -->

## How to install

1. Create a new Excel workbook
2. Press <kbd>ALT</kbd>-<kbd>F11</kbd> to open the Visual Basic Editor
3. Create a new module

    ![Insert a new module](./images/insert_module.png)

4. Copy/paste there the VBA code you can find below or in the `files/modWebService.bas` file.
5. Take a look to the declaration of the `InputXmlFile` constant: update the path to any valid path on your system and create that file.
6. Open that file and copy/paste there the content of the `files/checkVat.xml`

## How to call

Go in the Visual Basic Editor, put the cursor in the `run` subroutine and press <kbd>F5</kbd>. The code will retrieve the company behind the provided VAT number.

## Source

### modWebService.bas

```VBNet
<!-- concat-md::include "./files/modWebService.bas" -->
```

### checkVat.xml

```xml
<!-- concat-md::include "./files/checkVat.xml" -->
```
