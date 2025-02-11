# SIMPLE-XLSX

This project have scope simplify use plugin php spreadsheet.

This project create class for management creation file xlsx; into this class it's possible define:

- plus sheets
- header difference by number sheets
- possibility custom font color and background single colunm
- possibility custom define (if single sheet) define name single sheet
- possibility added before table, of information additional
- possibility create, after creation file xlsx, the zip

***
***

## Use

Into directory *example*, they are found example code. In general, from user respect this sequence:

1. instance class
<pre>$xlsx = new \Anton\SimpleXlsx\SimpleXlsx($header,'standard',null,1,null,$pathBase,null,null);</pre>

2. instane headers and sheets
<pre> $row  = $xlsx->setSpreadsheet();</pre>
This define initial base row.

3. Read your data and call method for create body:
<pre> $xlsx->setBodyCell(0,0,$row,$item['name'],$color);</pre>

4. Save file
<pre>$xlsx->save();</pre>

***

## DETAILS OF ARGUMENTS FROM THE INSTANCE CLASS AND FROM THE PRINCIPAL FUNCTION INTO CLASS

### INSTANCE

1. **header**: array header file; this array it must be how example:
<pre>$headeres = [ [] ]</pre>

2. **title** : name file
3. **sheets**: array one-dimensional and value into this it must be string
4. **default row**: this define where it begins table. Default is defined to 1; if greater one, table after the row defined
5. **extra data**: this can to be string or array; if is defined and defaultRow greater one, into the  file xlsx is positioned before table
6. **pathbase**: is directory of save file
7. **len**: if defined extradata this represents the columns length
8. **extracolor*: array, represents the possibility change color font and background

***

### FUNCTION

*setSpreadsheet*

This function create header and sheets.

The arguments are:

1. background: you can to be null, if is defined represents color of background the columns from header
2. color: you can to be null, if is defined represents color of font the columns from header
3. name sheet: you can to be null, if is defined only sheet and represents name of sheet


*setBodyCell*

With this function is created column from the body.

The arguments are:

- index sheet 
- index column (if zero example A0,B0,ec..)
- row 
- data
- boolean value: if true define background column difference ('F2F3F4' or 'EAEDED')
- position text : left,center or rigth
- number format data: for this argument it is postponed [NumberFormat](PhpOffice\PhpSpreadsheet\Style\NumberFormat)
- fill: the possibility custom background color column and bold or not bold text; this is an array

*save*

This function create file. It's possible call this function with argument ZipArchive; if this is defined the function create zip with file xlsx.




## EXAMPLE

- [STAMDARD](example/standard.php)
- [STANDARDWITH  EXTRA  DATA](example/standard  with  extra  data.php)
- [STANDARD  WITH  EXTRA  DATA  CUSTOM](example/standard  with  extra  data  custom.php)
- [STANDARD  WITH  ZIP](example/standard  with  zip.php)
- [TWO SHEETS](example/twoSheets.php)
- [HEADER  CUSTOM](example/header  custom.php)
- [FILL](example/fill.php)
- [SHEET  NAME  CUSTOM](example/sheet  name  custom.php)

