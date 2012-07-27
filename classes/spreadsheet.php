<?php defined('SYSPATH') or die('No direct access allowed.');
/**
 * PHP Excel library. Helper class to make and read spreadsheet easier
 * 
 * @package Koahana
 * @category spreadsheet
 * @author Katan
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * 
 * @see https://github.com/rafsoaken/kohana-phpexcel (Flynsarmy, Dmitry Shovchko)
 * 
 */
class Spreadsheet {
          
          /**
           * @var PHPExcel
           */
          protected $_spreadsheet;
          
          /**
           * @var array Valid types for PHPExcel
           */
          protected $options = array(
                'title' => 'New Spreadsheet',
                'subject' => 'New Spreadsheet',
                'description' => 'New Spreadsheet',
                'author' => 'None',
                'format' => 'Excel2007',
                'path' => '/',
                'name' => 'NewSpreadsheet',
                'filename' => '', // Filename for read
          );
          private $exts = array(
                'CSV' => 'csv',
                'PDF' => 'pdf',
                'Excel5' => 'xls',
                'Excel2007' => 'xlsx',
          );
          
          private $mimes = array(
                'CSV' => 'text/csv',
                'PDF' => 'application/pdf',
                'Excel5' => 'application/vnd.ms-excel',
                'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          );
          
          /**
           * Creates the spreadsheet class with given or default settings
           * @param array $options with optional parameters: title, subject, description, author
           * @param boolean $write default true, false to read filename (csv, xls, xlsx)
           * @return Spreadsheet 
           */
          public static function factory($options = array(), $write = TRUE)
          {
                    return new Spreadsheet($options, $write);
          }
          
          /**
           * 
           * @param array $options with optional parameters: title, subject, description, author
           */
          protected function __construct( Array $options, $write = TRUE)
          {
                    if ($write)
                    {
                              $this->_spreadsheet = new PHPExcel();
                              $this->set_options($options);
                    }
                    else
                    {
                              // Auto-identify format file
                              $options['format'] = PHPExcel_IOFactory::identify($options['filename']);
                              $this->set_options($options);
                    }
          }
          
          /**
           * Add/Update options
           * @param Array $options
           * @return Spreadsheet 
           */
          protected function set_options( Array $options)
          {
                    $this->options = Arr::merge($this->options, $options);
                    return $this;
          }
          
          public function get_options()
          {
                    return $this->options;
          }    
          
          /**
           * Creates a PHPExcel instance to load document for read
           * @param array $csv_values define delimiter and line ending
           * @return object PHPExcel 
           */
          public function load( $csv_values = Array('delimiter' => ';', 'lineEnding' => "\r\n"))
          {
                    // Create a new Reader defined in options
                    $this->_spreadsheet = PHPExcel_IOFactory::createReader($this->options['format']);
                    switch ($this->options['format'])
                    {
                              case 'CSV':
                                        $this->_spreadsheet->setDelimiter($csv_values['delimiter']);
                                        //Spreadsheet::$_spreadsheet->setEnclosure('');
                                        $this->_spreadsheet->setLineEnding($csv_values['lineEnding']);
                                        break;
                              case 'Excel2007' OR 'Excel5':
                                        $this->_spreadsheet->setReadDataOnly(true);
                                        break;
                    }
                    return $this;
          }
          
          /**
           * Return Array with all data from the first spreadsheet
           * @param array $valuetypes
           * @param array $skip content only numbers
           * @param boolean $emptyvalues
           * @return array 
           * 
           * @example
           * skip is an array with content only numbers, 1, 2, 3 ...
           * $emptyvalues remove arrays only if the all cells are empty
           */
          public function read($valuetypes = Array(), $skip = Array(), $emptyvalues = FALSE)
          {
                    /**
                     *@var array $array_data save parsed data from spreadsheet
                     */
                    $array_data = array();
                    foreach ($this->_spreadsheet->load($this->options['filename'])->getActiveSheet()->getRowIterator() as $i => $row)
                    {
                              
                              $cellIterator = $row->getCellIterator();
                              //skip rows in array
                              if ( ! empty($skip) AND in_array($i, $skip)) continue;
                              
                              //if ($skip[$i] == $row->getRowIndex()) continue;
                              $rowIndex = $row->getRowIndex();
                              $values = array();
                              
                              /** 
                               * @var PHPExcel_Cell $cell
                               */
                              foreach ($cellIterator as $cell) {
                                        if ( ! empty($valuetypes) AND array_key_exists($cell->getColumn(), $valuetypes))
                                        {
                                                  $format = explode(':', $valuetypes[$cell->getColumn()]);
                                                  switch ($format[0])
                                                  {
                                                            case 'date' : 
                                                                      $date = PHPExcel_Shared_Date::ExcelToPHPObject($cell->getValue());
                                                                      $array_data[$rowIndex][$cell->getColumn()] = $date->format($format[1]);
                                                                      break;
                                                  }
                                        }
                                        else
                                        {
                                                  // check if is_null or empty
                                                  $value = $cell->getValue();
                                                  $array_data[$rowIndex][$cell->getColumn()] = (strtolower($value) == 'null' OR empty($value))
                                                          ? null
                                                          : $cell->getCalculatedValue();
                                        }
                                        // For check empty values
                                        $values[] = $cell->getValue();
                              }
                              // Remove rows with all empty cells
                              if ($emptyvalues)
                              {
                                        $chechvalues = implode('', $values);
                                        if (empty($chechvalues))
                                        {
                                                  // Delete last array with empty values
                                                  array_pop($array_data);
                                        }
                              }
                    }
                    return (Array)$array_data;
          }
          
          /**
           * 
           * @return Spreadsheet 
           */
          protected function set_properties()
          {
                    $this->_spreadsheet->getProperties()
                            ->setCreator($this->options['author'])
                            ->setTitle($this->options['title'])
                            ->setSubject($this->options['subject'])
                            ->setDescription($this->options['description']);
                    return $this;
          }
          
          /**
           * Set active sheet index
           * 
           * @param int $index Active sheet index
           * @return void
           */
          public function set_active_sheet($index)
          {
                    $this->_spreadsheet->setActiveSheetIndex($index);
          }
          
          /**
           * Get the currently active sheet
           * 
           * @return PHPExcel_Worksheet
           */
          public function get_active_sheet()
          {
                    return $this->_spreadsheet->getActiveSheet();
          }

          /**
           * Writes cells to the spreadsheet
           * array(
           * 	   1 => array('A1', 'B1', 'C1', 'D1', 'E1'),
           * 	   2 => array('A2', 'B2', 'C2', 'D2', 'E2'),
           * 	   3 => array('A3', 'B3', 'C3', 'D3', 'E3'),
           * );
           * 
           * @param array of array( [row] => array([col]=>[value]) ) ie $arr[row][col] => value
           * @param boolean $multi_sheet for two or more sheets
           * @return void
           */
          public function set_data(array $data, $multi_sheet=false)
          {
                    //Single sheet ones can just dump everything to the current sheet
                    if ( !$multi_sheet )
                    {
                              $sheet = $this->_spreadsheet->getActiveSheet();
                              $this->set_sheet_data($data, $sheet);
                    }
                    //Have to do a little more work with multi-sheet
                    else
                    {
                              foreach ($data as $sheetname=>$sheetData)
                              {
                                        $sheet = $this->_spreadsheet->createSheet();
                                        $sheet->setTitle($sheetname);
                                        $this->set_sheet_data($sheetData, $sheet);
                              }
                              //Now remove the auto-created blank sheet at start of XLS
                              $this->_spreadsheet->removeSheetByIndex(0);
                    }
          }

          protected function set_sheet_data(array $data, PHPExcel_Worksheet $sheet)
          {
                    foreach ($data as $row =>$columns)
                              foreach ($columns as $column=>$value)
                                        $sheet->setCellValueByColumnAndRow($column, $row, $value);
          }

          /**
           * Writes spreadsheet to file
           * 
           * @return Path to spreadsheet
          */
          public function save()
          {
                    // Set document properties
                    $this->set_properties();
                    $writer = PHPExcel_IOFactory::createWriter($this->_spreadsheet, $this->options['format']);
                    
                    //Generate full path
                    $fullpath = $this->options['path'].$this->options['name'].'.'.$this->exts[$this->options['format']];
                    
                    if ($this->options['format'] == 'CSV')
                    {
                              $writer->setUseBOM(true);
                    }
                    $writer->save($fullpath);
                    return $fullpath;
          }

          /**
           * Send spreadsheet to browser without save to a file
           * @return void 
           */
          public function send()
          {                   
                    $response = Request::current()->response();
                    $response->send_file(
                            $this->save(),
                            $this->options['name'].'.'.$this->exts[$this->options['format']], // filename
                            array(
                                  'mime_type' => $this->mimes[$this->options['format']]
                            ));
          }
}
