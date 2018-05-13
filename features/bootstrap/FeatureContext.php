<?php

use Behat\Behat\Context\Context;
use Behat\Gherkin\Node\PyStringNode;
use Behat\Gherkin\Node\TableNode;
use Behat\Behat\Tester\Exception\PendingException;
use SpreadsheetExcelReader\Reader;

/**
 * Defines application features from the specific context.
 */
class FeatureContext implements Context
{
    /**
     * @var Spreadsheet_Excel_Reader
     */
    protected $excel;
    
    /**
     * Initializes context.
     *
     * Every scenario gets its own context instance.
     * You can also pass arbitrary arguments to the
     * context constructor through behat.yml.
     */
    public function __construct()
    {
    }
    
    /**
     * @Given document called :fileName
     */
    public function ilDocumentoChiamato($fileName)
    {
        $this->excel = new Reader();
        $this->excel->read(__DIR__ . '/../Fixtures/' . $fileName);
    }

    /**
     * @Then va bene così
     */
    public function vaBeneCosi()
    {
        
    }
}
