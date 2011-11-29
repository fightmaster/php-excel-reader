<?php
namespace PhpExcelReader;
define('IDENTIFIER_OLE', pack("CCCCCCCC",0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1));
class OLERead
{
    const NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
    const SMALL_BLOCK_DEPOT_BLOCK_POS    = 0x3c;
    const ROOT_START_BLOCK_POS           = 0x30;
    const BIG_BLOCK_SIZE                 = 0x200;
    const SMALL_BLOCK_SIZE               = 0x40;
    const EXTENSION_BLOCK_POS            = 0x44;
    const NUM_EXTENSION_BLOCK_POS        = 0x48;
    const PROPERTY_STORAGE_BLOCK_SIZE    = 0x80;
    const BIG_BLOCK_DEPOT_BLOCKS_POS     = 0x4c;
    const SMALL_BLOCK_THRESHOLD          = 0x1000;
    // property storage offsets
    const SIZE_OF_NAME_POS               = 0x40;
    const TYPE_POS                       = 0x42;
    const START_BLOCK_POS                = 0x74;
    const SIZE_POS                       = 0x78;

    public $error;
    public $numBigBlockDepotBlocks;
    public $sbdStartBlock;
    public $rootStartBlock;
    public $extensionBlock;
    public $numExtensionBlocks;
    public $bigBlockChain;
    public $smallBlockChain;
    public $entry;
    public $props;
    public $wrkbook;
    public $rootentry;

    public $data = '';

    private function GetInt4d($data, $pos) {
    	$value = ord($data[$pos]) | (ord($data[$pos+1])	<< 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);
    	if ($value>=4294967294) {
    		$value=-2;
    	}
    	return $value;
    }

    function read($sFileName)
    {
        // check if file exist and is readable (Darko Miljanovic)
        if(!is_readable($sFileName)) {
            $this->error = 1;
            return false;
        }
        $this->data = @file_get_contents($sFileName);
        if (!$this->data) {
            $this->error = 1;
            return false;
        }
        if (substr($this->data, 0, 8) != IDENTIFIER_OLE) {
            $this->error = 1;
            return false;
        }
        $this->numBigBlockDepotBlocks = $this->GetInt4d($this->data, self::NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
        $this->sbdStartBlock = $this->GetInt4d($this->data, self::SMALL_BLOCK_DEPOT_BLOCK_POS);
        $this->rootStartBlock = $this->GetInt4d($this->data, self::ROOT_START_BLOCK_POS);
        $this->extensionBlock = $this->GetInt4d($this->data, self::EXTENSION_BLOCK_POS);
        $this->numExtensionBlocks = $this->GetInt4d($this->data, self::NUM_EXTENSION_BLOCK_POS);

        $bigBlockDepotBlocks = array();
        $pos = self::BIG_BLOCK_DEPOT_BLOCKS_POS;
        $bbdBlocks = $this->numBigBlockDepotBlocks;
        if ($this->numExtensionBlocks != 0) {
            $bbdBlocks = (self::BIG_BLOCK_SIZE - self::BIG_BLOCK_DEPOT_BLOCKS_POS)/4;
        }

        for ($i = 0; $i < $bbdBlocks; $i++) {
            $bigBlockDepotBlocks[$i] = $this->GetInt4d($this->data, $pos);
            $pos += 4;
        }


        for ($j = 0; $j < $this->numExtensionBlocks; $j++) {
            $pos = ($this->extensionBlock + 1) * self::BIG_BLOCK_SIZE;
            $blocksToRead = min($this->numBigBlockDepotBlocks - $bbdBlocks, self::BIG_BLOCK_SIZE / 4 - 1);

            for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i++) {
                $bigBlockDepotBlocks[$i] = $this->GetInt4d($this->data, $pos);
                $pos += 4;
            }

            $bbdBlocks += $blocksToRead;
            if ($bbdBlocks < $this->numBigBlockDepotBlocks) {
                $this->extensionBlock = $this->GetInt4d($this->data, $pos);
            }
        }

        // readBigBlockDepot
        $index = 0;
        $this->bigBlockChain = array();

        for ($i = 0; $i < $this->numBigBlockDepotBlocks; $i++) {
            $pos = ($bigBlockDepotBlocks[$i] + 1) * self::BIG_BLOCK_SIZE;
            //echo "pos = $pos";
            for ($j = 0 ; $j < self::BIG_BLOCK_SIZE / 4; $j++) {
                $this->bigBlockChain[$index] = $this->GetInt4d($this->data, $pos);
                $pos += 4 ;
                $index++;
            }
        }

        // readSmallBlockDepot();
        $index = 0;
        $sbdBlock = $this->sbdStartBlock;
        $this->smallBlockChain = array();

        while ($sbdBlock != -2) {
          $pos = ($sbdBlock + 1) * self::BIG_BLOCK_SIZE;
          for ($j = 0; $j < self::BIG_BLOCK_SIZE / 4; $j++) {
            $this->smallBlockChain[$index] = $this->GetInt4d($this->data, $pos);
            $pos += 4;
            $index++;
          }
          $sbdBlock = $this->bigBlockChain[$sbdBlock];
        }


        // readData(rootStartBlock)
        $block = $this->rootStartBlock;
        $this->entry = $this->__readData($block);
        $this->__readPropertySets();
    }

    function __readData($bl) {
        $block = $bl;
        $data = '';
        while ($block != -2)  {
            $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
            $data = $data.substr($this->data, $pos, self::BIG_BLOCK_SIZE);
            $block = $this->bigBlockChain[$block];
        }
        return $data;
     }

    function __readPropertySets(){
        $offset = 0;
        while ($offset < strlen($this->entry)) {
            $d = substr($this->entry, $offset, self::PROPERTY_STORAGE_BLOCK_SIZE);
            $nameSize = ord($d[self::SIZE_OF_NAME_POS]) | (ord($d[self::SIZE_OF_NAME_POS+1]) << 8);
            $type = ord($d[self::TYPE_POS]);
            $startBlock = $this->GetInt4d($d, self::START_BLOCK_POS);
            $size = $this->GetInt4d($d, self::SIZE_POS);
            $name = '';
            for ($i = 0; $i < $nameSize ; $i++) {
                $name .= $d[$i];
            }
            $name = str_replace("\x00", "", $name);
            $this->props[] = array (
                'name' => $name,
                'type' => $type,
                'startBlock' => $startBlock,
                'size' => $size);
            if ((strtolower($name) == "workbook") || ( strtolower($name) == "book")) {
                $this->wrkbook = count($this->props) - 1;
            }
            if ($name == "Root Entry") {
                $this->rootentry = count($this->props) - 1;
            }
            $offset += self::PROPERTY_STORAGE_BLOCK_SIZE;
        }

    }


    function getWorkBook()
    {
        if ($this->props[$this->wrkbook]['size'] < self::SMALL_BLOCK_THRESHOLD){
            $rootdata = $this->__readData($this->props[$this->rootentry]['startBlock']);
            $streamData = '';
            $block = $this->props[$this->wrkbook]['startBlock'];
            while ($block != -2) {
                  $pos = $block * self::SMALL_BLOCK_SIZE;
                  $streamData .= substr($rootdata, $pos, self::SMALL_BLOCK_SIZE);
                  $block = $this->smallBlockChain[$block];
            }
            return $streamData;
        }else{
            $numBlocks = $this->props[$this->wrkbook]['size'] / self::BIG_BLOCK_SIZE;
            if ($this->props[$this->wrkbook]['size'] % self::BIG_BLOCK_SIZE != 0) {
                $numBlocks++;
            }

            if ($numBlocks == 0) return '';
            $streamData = '';
            $block = $this->props[$this->wrkbook]['startBlock'];
            while ($block != -2) {
              $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
              $streamData .= substr($this->data, $pos, self::BIG_BLOCK_SIZE);
              $block = $this->bigBlockChain[$block];
            }
            return $streamData;
        }
    }

}