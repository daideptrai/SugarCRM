<?php
if(!defined('sugarEntry') || !sugarEntry) die('Not A Valid Entry Point');

/*********************************************************************************
 * SugarCRM Community Edition is a customer relationship management program developed by
 * SugarCRM, Inc. Copyright (C) 2004-2013 SugarCRM Inc.
 * 
 * This program is free software; you can redistribute it and/or modify it under
 * the terms of the GNU Affero General Public License version 3 as published by the
 * Free Software Foundation with the addition of the following permission added
 * to Section 15 as permitted in Section 7(a): FOR ANY PART OF THE COVERED WORK
 * IN WHICH THE COPYRIGHT IS OWNED BY SUGARCRM, SUGARCRM DISCLAIMS THE WARRANTY
 * OF NON INFRINGEMENT OF THIRD PARTY RIGHTS.
 * 
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE.  See the GNU Affero General Public License for more
 * details.
 * 
 * You should have received a copy of the GNU Affero General Public License along with
 * this program; if not, see http://www.gnu.org/licenses or write to the Free
 * Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301 USA.
 * 
 * You can contact SugarCRM, Inc. headquarters at 10050 North Wolfe Road,
 * SW2-130, Cupertino, CA 95014, USA. or at email address contact@sugarcrm.com.
 * 
 * The interactive user interfaces in modified source and object code versions
 * of this program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU Affero General Public License version 3.
 * 
 * In accordance with Section 7(b) of the GNU Affero General Public License version 3,
 * these Appropriate Legal Notices must retain the display of the "Powered by
 * SugarCRM" logo. If the display of the logo is not reasonably feasible for
 * technical reasons, the Appropriate Legal Notices must display the words
 * "Powered by SugarCRM".
 ********************************************************************************/

/*********************************************************************************

 * Description: Class to handle processing an import file
 * Portions created by SugarCRM are Copyright (C) SugarCRM, Inc.
 * All Rights Reserved.
 ********************************************************************************/

require_once('modules/Import/CsvAutoDetect.php');
require_once('modules/Import/sources/ImportDataSource.php');
require_once('custom/include/PHPExcel1.8.0/PHPExcel.php');

class ImportFileXls extends ImportDataSource
{
    /**
     * Stores whether or not we are deleting the import file in the destructor
     */
    private $_deleteFile;

    /**
     * File pointer returned from fopen() call
     */
    private $_fp = FALSE;

    /**
     * True if the csv file has a header row.
     */
    private $_hasHeader = FALSE;

    /**
     * True if the csv file has a header row.
     */
    private $_detector = null;

    /**
     * CSV date format
     */
    private $_date_format = false;

    /**
     * CSV time format
     */
    private $_time_format = false;

    /**
     * The import file map that this import file inherits properties from.
     */
    private $_importFile = null;

    /**
     * Delimiter string we are using (i.e. , or ;)
     */
    private $_delimiter;

    /**
     * Enclosure string we are using (i.e. ' or ")
     */
    private $_enclosure;
    
    /**
     * File encoding, used to translate the data into UTF-8 for display and import
     */
    private $_encoding;
    
    private $_objWorksheet = false;
    private $_maxCol = 0;


    /**
     * Constructor
     *
     * @param string $filename
     * @param string $delimiter
     * @param string $enclosure
     * @param bool   $deleteFile
     */
    public function __construct( $filename, $delimiter  = ',', $enclosure  = '',$deleteFile = true, $checkUploadPath = TRUE )
    {
        if ( !is_file($filename) || !is_readable($filename) ) {
            return false;
        }

        if ( $checkUploadPath && UploadStream::path($filename) == null )
        {
            $GLOBALS['log']->fatal("ImportFile detected attempt to access to the following file not within the sugar upload dir: $filename");
            return null;
        }

        // turn on auto-detection of line endings to fix bug #10770
        ini_set('auto_detect_line_endings', '1');

        //$this->_fp         = sugar_fopen($filename,'r');
        $this->_fp = true;
        $this->_sourcename   = $filename;
        $this->_deleteFile = $deleteFile;
        
        $type = 'Excel5';
        $objReader   = \PHPExcel_IOFactory::createReader($type);
        $objPHPExcel = $objReader->load($this->_sourcename);
        
        $this->_objWorksheet = $objPHPExcel->setActiveSheetIndex(0);
        $this->_maxCol = count($this->_objWorksheet->getColumnDimensions());
        
        $this->_rowsCount = 1;
        //$this->setHeaderRow($this->getNextRow());
        /*
        $this->_delimiter  = ( empty($delimiter) ? ',' : $delimiter );
        if ($this->_delimiter == '\t') {
            $this->_delimiter = "\t";
        }
        $this->_enclosure  = ( empty($enclosure) ? '' : trim($enclosure) );

        // Autodetect does setFpAfterBOM()
        $this->_encoding = $this->autoDetectCharacterSet();*/
    }

    /**
     * Remove the BOM (Byte Order Mark) from the beginning of the import row if it exists
     * @return void
     */
    private function setFpAfterBOM()
    {
        return;
    }
    /**
     * Destructor
     *
     * Deletes $_importFile if $_deleteFile is true
     */
    public function __destruct()
    {
        if ( $this->_deleteFile && $this->fileExists() ) {
            fclose($this->_fp);
            //Make sure the file exists before unlinking
            if(file_exists($this->_sourcename)) {
               unlink($this->_sourcename);
            }
        }

        ini_restore('auto_detect_line_endings');
    }

    /**
	 * This is needed to prevent unserialize vulnerability
     */
    public function __wakeup()
    {
        // clean all properties
        foreach(get_object_vars($this) as $k => $v) {
            $this->$k = null;
        }
        throw new Exception("Not a serializable object");
    }

    /**
     * Returns true if the filename given exists and is readable
     *
     * @return bool
     */
    public function fileExists()
    {
    	return !$this->_fp ? false : true;
    }

    /**
     * Gets the next row from $_importFile
     *
     * @return array current row of file
     */
    public function getNextRow()
    {   
        if ($this->_rowsCount > $this->getNumberOfLinesInfile()) {
            return false;
        }
        
        $this->_currentRow = array();
        
        for ($i = 0; $i < $this->_maxCol; $i++)
        {
            $this->_currentRow[] = $this->_objWorksheet->getCellByColumnAndRow($i, $this->_rowsCount)->getValue();
        }
        
        $this->_rowsCount++;
        
        return $this->_currentRow;
    }

    /**
     * Returns the number of fields in the current row
     *
     * @return int count of fiels in the current row
     */
    public function getFieldCount()
    {
        return count($this->_currentRow);
    }

    /**
     * Determine the number of lines in this file.
     *
     * @return int
     */
    public function getNumberOfLinesInfile()
    {
        return $this->_objWorksheet->getHighestRow();
    }

    //TODO: Add auto detection for field delim and qualifier properteis.
    public function autoDetectCSVProperties()
    {
        // defaults
        $this->_delimiter  = ",";
        $this->_enclosure  = '"';

        $this->_detector = new CsvAutoDetect($this->_sourcename);

        $delimiter = $enclosure = false;

        $ret = $this->_detector->getCsvSettings($delimiter, $enclosure);
        if ($ret)
        {
            $this->_delimiter = $delimiter;
            $this->_enclosure = $enclosure;
            return TRUE;
        }
        else
        {
            return FALSE;
        }
    }

    public function getFieldDelimeter()
    {
        return $this->_delimiter;
    }

    public function getFieldEnclosure()
    {
        return $this->_enclosure;
    }

    public function autoDetectCharacterSet()
    {
        return 'UTF-8';
    }

    public function getDateFormat()
    {
        if ($this->_detector) {
            $this->_date_format = $this->_detector->getDateFormat();
        }

        return $this->_date_format;
    }

    public function getTimeFormat()
    {
        if ($this->_detector) {
            $this->_time_format = $this->_detector->getTimeFormat();
        }

        return $this->_time_format;
    }

    public function setHeaderRow($hasHeader)
    {
        $this->_hasHeader = $hasHeader;
    }

    public function hasHeaderRow($autoDetect = TRUE)
    {
        $this->_hasHeader = true;
        return $this->_hasHeader;
    }

    public function setImportFileMap($map)
    {
        $this->_importFile = $map;
        $importMapProperties = array('_delimiter' => 'delimiter','_enclosure' => 'enclosure', '_hasHeader' => 'has_header');
        //Inject properties from the import map
        foreach($importMapProperties as $k => $v)
        {
            $this->$k = $map->$v;
        }
    }

    //Begin Implementation for SPL's Iterator interface
    public function key()
    {
        return $this->_rowsCount;
    }

    public function current()
    {
        return $this->_currentRow;
    }

    public function next()
    {
        $this->getNextRow();
    }

    public function valid()
    {
        return $this->_currentRow !== FALSE;
    }

    public function rewind()
    {
        $this->setFpAfterBOM();
        //Load our first row
        $this->getNextRow();
    }

    public function getTotalRecordCount()
    {
        $totalCount = $this->getNumberOfLinesInfile();
        if($this->hasHeaderRow(FALSE) && $totalCount > 0)
        {
            $totalCount--;
        }
        return $totalCount;
    }

    public function loadDataSet($totalItems = 0)
    {
        $currentLine = 0;
        $this->_dataSet = array();
        $this->rewind();
        //If there's a header don't include it.
        if( $this->hasHeaderRow(FALSE) )
            $this->next();

        while( $this->valid() &&  $totalItems > count($this->_dataSet) )
        {
            if($currentLine >= $this->_offset)
            {
                $this->_dataSet[] = $this->_currentRow;
            }
            $this->next();
            $currentLine++;
        }

        return $this;
    }

    public function getHeaderColumns()
    {
        $this->rewind();
        if($this->hasHeaderRow(FALSE))
            return $this->_currentRow;
        else
            return FALSE;
    }

}
