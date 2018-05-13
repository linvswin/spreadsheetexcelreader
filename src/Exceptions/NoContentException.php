<?php

namespace SpreadsheetExcelReader\Exceptions;

class NoContentException extends \Exception{
    public static function default(){
        return new self('File has no content');
    }
}
