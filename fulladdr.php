<?php
//
// generate some random data and add in some spreadsheet functions which will be rendered when the 
// file is loaded into a spreadsheet program such as excel desktop & online as well as google sheets
// it's just a demonstration of a technique

//

$delim = "\t";
$rowsNeeded = 10;
$colsNeeded = 4;
$colsNeededWithAverage = $colsNeeded + 1;
$delimBuffer = "";

showHeader($colsNeeded,$delim);
//
// generate the data
//
showData($rowsNeeded,$colsNeeded,$delim);
//
// we've displayed the base data, now the subtotals
//
// put a blank line in to make selecting the data easier in the spreadsheet
//
echo PHP_EOL;

subtotals($colsNeededWithAverage,$delim);
//
// all done
//

function showHeader($colsNeeded,$delim) {
    //
    // create the spereadsheet column names header
    //
    // we know how many columns are needed.....
    //
    $hdr = array();
    for ($col = 1;$col <= $colsNeeded;$col++) {
        $hdr[] = "number" . $col;
    }
    //
    // .... plus one for the average column
    //
    $hdr[] = "average";
    $showHdr = implode($delim,$hdr);
    echo $showHdr . PHP_EOL;
}

function showData($rowsNeeded,$colsNeeded,$delim) { 

    $sheetRow = array();    

    for($row = 1;$row<=$rowsNeeded;$row++) {  
        
        for($col = 1;$col <= $colsNeeded;$col++){
            $number = rand(1,5000);                   
            $sheetRow[] = $number;
        }

        $ref1 = coordsToLabel(1,$row+1);
        $ref2 = coordsToLabel($colsNeeded,$row+1);
    //
    // we could calculate the average here rather than use excel but this would make changing values difficult at run-time as it would
    // be hard coded, so let the spreadsheet program do it.
    // As an example of having more than one function in the cell, We build up the function to get the average then wrap that with
    // int to get just the integer part. You can add as many functions as you like here.
    // Finaly put an = at the start, note you only use a = at the very start not on each function.
    //
        $exe = "average($ref1:$ref2)";
        $exe = "int(" . $exe . ")";
        $exe = "=" . $exe;
        
        $sheetRow[] = $exe;
        $showRow = implode($delim,$sheetRow);
        echo $showRow;     
        echo PHP_EOL; 
        unset ($sheetRow);   
    }
}

function subtotals($colsNeededWithAverage,$delim) {
//
// When subtotals are used (in excel anyway) the =subtotal works on visible rows, I've just used a filter (data->filter in the excel UI) 
// but you need to msake the sure the =subtotal itself is not in the filter set as it will be evaluated as well and thats probabbly 
// not what is required. 
//
// When I've tested this code, I just selected the data by draging frpom the bottomn right corner to thye4 top left then do the 
// data->filter on that selection. Or you could display the =subtotal to the right away from the main data and just choose the
// columns for the data->filter which might be qukcker if you have many rows, just make sure the =subtotals are not in the selected set.
// I simply generated a string with delimiters in which makes blank columns. It's towards the end of the function
//
// An issue I've seen with subtotal() and probabbly other functions in online Excel is it needed a semi-colon rather than a comma 
// after the 109 (in this case).Seems my language was set to French and so needed a semi-colon. Set languae to UK and comma worked. Desktop
// behaves the same way. Google sheets works either way.
//
    $sheetRow = array();
    $delimBuffer = "";

    for($col = 1;$col <= $colsNeededWithAverage;$col++) {  

        $ref1 = coordsToLabel($col,2);
        $ref2 = coordsToLabel($col,11);

        $exe = "subtotal(109,$ref1:$ref2)";
        $exe = "=" . $exe;
        $sheetRow[] = $exe;  
    }

    $showRow = implode($delim,$sheetRow); 

    for ($temp = 1;$temp <= $colsNeededWithAverage;$temp++) {
        $delimBuffer = $delimBuffer . $delim; 
    }

    //echo $delimBuffer; // uncomment this line if you want the =subtotals to the right for example
    echo $showRow;
 }

function coordsToLabel($column,$row,$relativeRequired=FALSE) { 
    //
    //  This function takes a numeric column number (e.g. as part of a loop) and calculates the letter needed in
    //  excel functions. Example is 1=A , 26=Z , 702=ZZ, 703=703 
    //  relative or absolute references can be generated using the $relativeRequired parameter
    // 
    //  in short $A$1 is an absolute reference and A1 is a relative reference
    //
    //  it is possible to refer to cells by R1C1 notation but in excel that needs to whole sheet to be configured to use those
    //  values rather than A1 notation.
    //
    //  it's not entirely straightforward to create the references , it's not just base 26. The article below reveals all
    // 
    //  adapted from https://learn.microsoft.com/en-us/office/troubleshoot/excel/convert-excel-column-numbers
    //
    // most spreadsheets have an address =function (and =row & =col) but these execute when the spreadsheet runs so not really practical
    // here as the address are references to other cells. You could use =row where I've used a row number and =col where the column is used
    // which ,might make life a little easier by just using $row and $col variables without the references, it is possible to use functions
    // like indirect to get the desired result I just fancied doing it this way 
    //

    $a = $column;
    $result = "";
    while ($column > 0) {
       $a = intval( ($column - 1) / 26);
       $b = ($column - 1) % 26;
       $result = chr($b + 65) . $result;
       $column = $a;
    }
 
    if ($relativeRequired) {
        $result = "\$" . $result . "\$" . $row;
    } else {
        $result = $result . $row;
    }
    return $result;
 }