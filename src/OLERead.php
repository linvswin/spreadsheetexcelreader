<?php
namespace SpreadsheetExcelReader;

use SpreadsheetExcelReader\Exceptions\NotReadableException;
use SpreadsheetExcelReader\Exceptions\NoContentException;
use SpreadsheetExcelReader\Exceptions\FileNotValidException;

/**
 * A class for reading Microsoft Excel Spreadsheets.
 *
 * Originally developed by Vadim Tkachenko under the name PHPExcelReader.
 * (http://sourceforge.net/projects/phpexcelreader)
 * Based on the Java version by Andy Khan (http://www.andykhan.com).  Now
 * maintained by David Sanders.  Reads only Biff 7 and Biff 8 formats.
 *
 * PHP versions 5
 *
 * LICENSE: This source file is subject to version 3.0 of the PHP license
 * that is available through the world-wide-web at the following URI:
 * http://www.php.net/license/3_0.txt.  If you did not receive a copy of
 * the PHP License and are unable to obtain it through the web, please
 * send a note to license@php.net so we can mail you a copy immediately.
 *
 * @category   Spreadsheet
 * @package    Spreadsheet_Excel_Reader
 * @author     Leonardo Fontana
 * @license    http://www.php.net/license/3_0.txt  PHP License 3.0
 * @version    CVS: $Id: reader.php 19 2007-03-13 12:42:41Z shangxiao $
 * @link       http://pear.php.net/package/Spreadsheet_Excel_Reader
 * @see        OLE, Spreadsheet_Excel_Writer
 */

class OLERead
{
    const NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
    const SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
    const ROOT_START_BLOCK_POS = 0x30;
    const BIG_BLOCK_SIZE = 0x200;
    const SMALL_BLOCK_SIZE = 0x40;
    const EXTENSION_BLOCK_POS = 0x44;
    const NUM_EXTENSION_BLOCK_POS = 0x48;
    const PROPERTY_STORAGE_BLOCK_SIZE = 0x80;
    const BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;
    const SMALL_BLOCK_THRESHOLD = 0x1000;
    // property storage offsets
    const SIZE_OF_NAME_POS = 0x40;
    const TYPE_POS = 0x42;
    const START_BLOCK_POS = 0x74;
    const SIZE_POS = 0x78;
    const MAX_IT_VALUE = 4294967294;

    protected static $IDENTIFIER_OLE = null;

    protected $data = '';


    public function __construct()
    {
        self::$IDENTIFIER_OLE = \pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
    }

    public static function GetInt4d( $data, $pos)
    {
        $value = ord($data[$pos]) | (ord($data[$pos + 1]) << 8) | (ord($data[$pos + 2]) << 16) | (ord($data[$pos + 3]) << 24);
        if ($value >= self::MAX_IT_VALUE)
            {
            $value = -2;
        }
        return $value;
    }

    public function read($sFileName)
    {
        
    	// check if file exist and is readable (Darko Miljanovic)
        if (!is_readable($sFileName)) {
            throw NotReadableException::default($sFileName);
        }

        $this->data = @file_get_contents($sFileName);
        if (!$this->data) {
            throw NoContentException::default($sFileName);
        }
        if (substr($this->data, 0, 8) != self::$IDENTIFIER_OLE) {
            throw new FileNotValidException($sFileName);
        }
        $this->numBigBlockDepotBlocks = self::GetInt4d($this->data, self::NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
        $this->sbdStartBlock = self::GetInt4d($this->data, self::SMALL_BLOCK_DEPOT_BLOCK_POS);
        $this->rootStartBlock = self::GetInt4d($this->data, self::ROOT_START_BLOCK_POS);
        $this->extensionBlock = self::GetInt4d($this->data, self::EXTENSION_BLOCK_POS);
        $this->numExtensionBlocks = self::GetInt4d($this->data, self::NUM_EXTENSION_BLOCK_POS);


        $bigBlockDepotBlocks = array();
        $pos = self::BIG_BLOCK_DEPOT_BLOCKS_POS;
        $bbdBlocks = $this->numBigBlockDepotBlocks;

        if ($this->numExtensionBlocks != 0) {
            $bbdBlocks = (self::BIG_BLOCK_SIZE - self::BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
        }

        for ($i = 0; $i < $bbdBlocks; $i++) {
            $bigBlockDepotBlocks[$i] = self::GetInt4d($this->data, $pos);
            $pos += 4;
        }


        for ($j = 0; $j < $this->numExtensionBlocks; $j++) {
            $pos = ($this->extensionBlock + 1) * self::BIG_BLOCK_SIZE;
            $blocksToRead = min($this->numBigBlockDepotBlocks - $bbdBlocks, self::BIG_BLOCK_SIZE / 4 - 1);

            for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i++) {
                $bigBlockDepotBlocks[$i] = self::GetInt4d($this->data, $pos);
                $pos += 4;
            }

            $bbdBlocks += $blocksToRead;
            if ($bbdBlocks < $this->numBigBlockDepotBlocks) {
                $this->extensionBlock = self::GetInt4d($this->data, $pos);
            }
        }

       // var_dump($bigBlockDepotBlocks);
        
        // readBigBlockDepot
        $pos = 0;
        $index = 0;
        $this->bigBlockChain = array();

        for ($i = 0; $i < $this->numBigBlockDepotBlocks; $i++) {
            $pos = ($bigBlockDepotBlocks[$i] + 1) * self::BIG_BLOCK_SIZE;
            for ($j = 0; $j < self::BIG_BLOCK_SIZE / 4; $j++) {
                $this->bigBlockChain[$index] = self::GetInt4d($this->data, $pos);
                $pos += 4;
                $index++;
            }
        }

        $pos = 0;
        $index = 0;
        $sbdBlock = $this->sbdStartBlock;
        $this->smallBlockChain = array();

        while ($sbdBlock != -2) {

            $pos = ($sbdBlock + 1) * self::BIG_BLOCK_SIZE;

            for ($j = 0; $j < self::BIG_BLOCK_SIZE / 4; $j++) {
                $this->smallBlockChain[$index] = self::GetInt4d($this->data, $pos);
                $pos += 4;
                $index++;
            }

            $sbdBlock = $this->bigBlockChain[$sbdBlock];
        }

        
        // readData(rootStartBlock)
        $block = $this->rootStartBlock;
        $pos = 0;
        $this->entry = $this->__readData($block);

        $this->__readPropertySets();

    }

    public function __readData($bl)
    {
        $block = $bl;
        $pos = 0;
        $data = '';

        while ($block != -2) {
            $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
            $data = $data . substr($this->data, $pos, self::BIG_BLOCK_SIZE);
            //echo "pos = $pos data=$data\n";	
            $block = $this->bigBlockChain[$block];
        }
        return $data;
    }

    public function __readPropertySets()
    {
        $offset = 0;
        while ($offset < strlen($this->entry)) {
            $d = substr($this->entry, $offset, self::PROPERTY_STORAGE_BLOCK_SIZE);

            $nameSize = ord($d[self::SIZE_OF_NAME_POS]) | (ord($d[self::SIZE_OF_NAME_POS + 1]) << 8);

            $type = ord($d[self::TYPE_POS]);

            $startBlock = self::GetInt4d($d, self::START_BLOCK_POS);
            $size = self::GetInt4d($d, self::SIZE_POS);

            $name = '';
            for ($i = 0; $i < $nameSize; $i++) {
                $name .= $d[$i];
            }

            $name = str_replace("\x00", "", $name);

            $this->props[] = array(
                'name' => $name,
                'type' => $type,
                'startBlock' => $startBlock,
                'size' => $size
            );

            if ( ($name == "Workbook") || ($name == "Book")) {
                $this->wrkbook = count($this->props) - 1;
            }

            if ($name == "Root Entry") {
                $this->rootentry = count($this->props) - 1;
            }

            $offset += self::PROPERTY_STORAGE_BLOCK_SIZE;
        }

    }


    public function getWorkBook()
    {
        if ($this->props[$this->wrkbook]['size'] < self::SMALL_BLOCK_THRESHOLD) {

            $rootdata = $this->__readData($this->props[$this->rootentry]['startBlock']);

            $streamData = '';
            $block = $this->props[$this->wrkbook]['startBlock'];
            $pos = 0;
            while ($block != -2) {
                $pos = $block * self::SMALL_BLOCK_SIZE;
                $streamData .= substr($rootdata, $pos, self::SMALL_BLOCK_SIZE);

                $block = $this->smallBlockChain[$block];
            }

            return $streamData;


        }
        else {

            $numBlocks = $this->props[$this->wrkbook]['size'] / self::BIG_BLOCK_SIZE;
            if ($this->props[$this->wrkbook]['size'] % self::BIG_BLOCK_SIZE != 0) {
                $numBlocks++;
            }

            if ($numBlocks == 0) return '';

            $streamData = '';
            $block = $this->props[$this->wrkbook]['startBlock'];
            $pos = 0;
            while ($block != -2) {
                $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
                $streamData .= substr($this->data, $pos, self::BIG_BLOCK_SIZE);
                $block = $this->bigBlockChain[$block];
            }
            return $streamData;
        }
    }

}

