<?php

use Behat\Behat\Context\Context;
use Behat\Gherkin\Node\PyStringNode;
use Behat\Gherkin\Node\TableNode;
use Behat\Behat\Tester\Exception\PendingException;
use Spreadsheet\Excel\Reader;
use PHPUnit\Framework\Assert;

/**
 * Defines application features from the specific context.
 */
class FeatureContext implements Context
{
    /**
     * @var Reader
     */
    protected $excel;

    /**
     * @var int
     */
    protected $sheet = 0;
    
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
    public function givenDocument($fileName)
    {
        $this->excel = new Reader();
        $this->excel->read(__DIR__ . '/../Fixtures/' . $fileName);
    }

    /**
     * @Given sheet :index
     */
    public function givenSheet($index){
        $this->sheet = $index;
    }

    /**
     * @Then the row :row and the column :col contains :expected
     */
    public function cellContains($row, $col, $expected)
    {
        $cells = $this->excel->sheets[$this->sheet]['cells'];
        Assert::arrayHasKey($row)->evaluate($cells, "the sheet doesn't contains row $row");
        $rowValue = $cells[$row];
        Assert::arrayHasKey($col)->evaluate($rowValue, "The row doesn't contain column $col");
        $value = $rowValue[$col];
        Assert::assertEquals($expected, $value);
    }

    /**
     * @Then the row :row and the column :col is empty
     */
    public function cellEmpty($row, $col){
        Assert::assertArrayNotHasKey($col, $this->excel->sheets[$this->sheet]['cells'][$row]);
    }

    /**
     *@Then the sheet exists 
    */
    public function sheetExists(){
        Assert::assertArrayHasKey($this->sheet, $this->excel->sheets);
    }
}
