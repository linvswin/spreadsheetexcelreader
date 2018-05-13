<?php

namespace SpreadsheetExcelReader\Exceptions;

class NotReadableException extends \Exception
{
    public static function throw($fileName){
        throw new self("$fileName is not readable");
    }
}
