<?php

namespace SpreadsheetExcelReader\Exceptions;

class FileNotValidException extends \Exception{
    public static function default(string $fileName): self{
        return new self("$fileName is not a valid Excel file");
    }
}