<?php

defined('SYSPATH') or die('No direct access allowed.');

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
class Kohana_Spreadsheet {

    /**
     * @var PHPExcel
     */
    protected $_spreadsheet;

    /**
     * @var object worksheet
     */
    protected $_worksheets = array();

    /**
     * @var array Valid types for PHPExcel
     */
    protected $options = array(
        'title' => 'New Spreadsheet',
        'subject' => 'New Spreadsheet',
        'description' => 'New Spreadsheet',
        'author' => 'None',
        'format' => 'Excel2007',
        'path' => './',
        'name' => 'NewSpreadsheet',
        'filename' => '', // Filename for read
        'csv_values' => array('delimiter' => ';', 'lineEnding' => "\r\n")// CSV file
    );

    /**
     * @var array file extentions
     */
    private $exts = array(
        'CSV' => 'csv',
        'PDF' => 'pdf',
        'Excel5' => 'xls',
        'Excel2007' => 'xlsx',
    );

    /**
     * @var array file mimes
     */
    private $mimes = array(
        'CSV' => 'text/csv',
        'PDF' => 'application/pdf',
        'Excel5' => 'application/vnd.ms-excel',
        'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    /**
     * Creates the spreadsheet class with given or default settings
     * @param array $options with optional parameters: title, subject, description, author
     * @return Spreadsheet 
     */
    public static function factory($options = array()) {

        return new Spreadsheet($options);
    }

    /**
     * 
     * @access protected
     * 
     * @param array $options with optional parameters: title, subject, description, author
     */
    protected function __construct(array $options) {

        //get PHPExcel instance
        $this->_spreadsheet = new PHPExcel();

        //load worksheets
        $this->load_worksheets();

        //set options
        $this->set_options($options);
    }

    /**
     * init worksheets
     * 
     * @access protected
     */
    protected function load_worksheets() {

        //empty the worksheets
        $this->_worksheets = array();

        foreach ($this->_spreadsheet->getWorksheetIterator() as $i => $worksheet) {

            //create default worksheet
            $worksheet = new Worksheet($this->_spreadsheet, $worksheet);
            $this->_worksheets[$i] = $worksheet;
        }

        return $this;
    }

    /**
     * create new worksheet
     * 
     * @access protected
     */
    protected function create_worksheet() {

        $newsheet = $this->_spreadsheet->createSheet();
        $new_worksheet = new Worksheet($this->_spreadsheet, $newsheet);
        $this->_worksheets[$this->_spreadsheet->getIndex($newsheet)] = $new_worksheet;

        return $new_worksheet;
    }

    /**
     * remove worksheet
     * 
     * @access protected
     * @param int $index worksheet index
     */
    protected function remove_worksheet($index) {

        $this->_spreadsheet->removeSheetByIndex($index);
        unset($this->_worksheets[$index]);
    }

    /**
     * Add/Update options
     * @param Array $options
     * @return Spreadsheet 
     */
    protected function set_options(array $options) {
        $this->options = Arr::merge($this->options, $options);
        return $this;
    }

    /**
     * Get options
     * 
     * @access public
     * @return array options
     */
    public function get_options() {
        return $this->options;
    }

    /**
     * Creates a PHPExcel instance to load document for read
     * @param array $csv_values define delimiter and line ending
     * @return object PHPExcel 
     */
    public function load($spreadsheet_file = NULL) {

        //if not specify spreadsheet file the load file from options
        if ($spreadsheet_file === NULL)
            $spreadsheet_file = $this->options['path'] . $this->options['filename'];

        // Auto-identify format file
        $this->options['format'] = PHPExcel_IOFactory::identify($spreadsheet_file);

        // Create a new Reader defined in options
        $spreadsheet_reader = PHPExcel_IOFactory::createReader($this->options['format']);

        switch ($this->options['format']) {
            case 'CSV':
                $spreadsheet_reader->setDelimiter($this->options['csv_values']['delimiter']);
                $spreadsheet_reader->setLineEnding($this->options['csv_values']['lineEnding']);
                break;
            case 'Excel2007' OR 'Excel5':
                $spreadsheet_reader->setReadDataOnly(true);
                break;
        }

        //load the spreadsheet
        $this->_spreadsheet = $spreadsheet_reader->load($spreadsheet_file);

        //reset worksheets
        $this->load_worksheets();

        //return
        return $this;
    }

    /**
     * Return Array with all data from the active spreadsheet
     * @param array $valuetypes
     * @param array $skip content only numbers
     * @param boolean $emptyvalues
     * @return array 
     * 
     * @example
     * skip is an array with content only numbers, 1, 2, 3 ...
     * $emptyvalues remove arrays only if the all cells are empty
     */
    public function read($valuetypes = array(), $skip = array(), $emptyvalues = FALSE) {
        return $this->get_active_worksheet()->read($valuetypes, $skip, $emptyvalues);
    }

    /**
     * 
     * @return Spreadsheet 
     */
    protected function set_properties() {

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
    public function set_active_worksheet($index) {
        $this->_spreadsheet->setActiveSheetIndex($index);
    }

    /**
     * Get the currently active sheet
     * 
     * @return PHPExcel_Worksheet
     */
    public function get_active_worksheet() {
        return $this->_worksheets[$this->_spreadsheet->getActiveSheetIndex()];
    }

    /**
     * Get one or more worksheets
     * 
     * @return mixed one more worksheets
     */
    public function get_worksheet($index = NULL) {

        if ($index === NULL)
            return $this->_worksheets;

        return $this->_worksheets[$index];
    }

    /**
     * call PHPExcel spreadsheet's function if not exist
     * 
     * @access public
     * @param string $method_name method name
     * @param mixed $arguments arguments
     */
    public function __call($method_name, $arguments) {

        $this->_spreadsheet->$method_name($arguments);
    }

    /**
     * Writes cells to the spreadsheet
     * array(
     *   'names'=>array('name1','name2','name4'),
     *   'rows'=>array(
     *      1 => array('A1', 'B1', 'C1', 'D1', 'E1'),
     *      2 => array('A2', 'B2', 'C2', 'D2', 'E2'),
     *      3 => array('A3', 'B3', 'C3', 'D3', 'E3'),
     * ));
     * 
     * @param array of array( [row] => array([col]=>[value]) ) ie $arr[row][col] => value
     * @param boolean $multi_sheet for two or more sheets
     * @return void
     */
    public function set_data(array $data, $multi_sheet = FALSE) {

        //Single sheet ones can just dump everything to the current sheet
        if (!$multi_sheet) {
            $worksheet = $this->get_active_worksheet();

            if (isset($data['columns']))
                $worksheet->columns($data['columns']);

            if (isset($data['types']))
                $worksheet->types($data['types']);

            if (isset($data['formats']))
                $worksheet->formats($data['formats']);

            $worksheet->data($data['rows']);
            $worksheet->render();
        }

        //Have to do a little more work with multi-sheet
        else {
            foreach ($data as $sheetname => $sheet_data) {

                $worksheet = $this->create_worksheet();
                $worksheet->title($sheetname);

                if (isset($sheet_data['columns']))
                    $worksheet->columns($sheet_data['columns']);

                if (isset($sheet_data['types']))
                    $worksheet->types($sheet_data['types']);

                if (isset($sheet_data['formats']))
                    $worksheet->formats($sheet_data['formats']);

                $worksheet->data($sheet_data['rows']);
                $worksheet->render();
            }

            //Now remove the auto-created blank sheet at start of XLS
            $this->remove_worksheet(0);
        }
    }

    /**
     * Writes spreadsheet to file
     * 
     * @return Path to spreadsheet
     */
    public function save() {

        // Set document properties
        $this->set_properties();
        $writer = PHPExcel_IOFactory::createWriter($this->_spreadsheet, $this->options['format']);

        //if 'path' not set, use temp dir instead
        if ($this->options['path'] === NULL)
            $this->options['path'] = sys_get_temp_dir() . DIRECTORY_SEPARATOR;

        //Generate full path
        $fullpath = $this->options['path'] . $this->options['name'] . '.' . $this->exts[$this->options['format']];

        if ($this->options['format'] == 'CSV') {
            $writer->setUseBOM(true);
        }
        $writer->save($fullpath);
        return $fullpath;
    }

    /**
     * Send spreadsheet to browser without save to a file
     * @return void 
     */
    public function send() {
        $response = Response::factory();
        $response->send_file(
                $this->save(), $this->options['name'] . '.' . $this->exts[$this->options['format']], // filename
                array(
            'mime_type' => $this->mimes[$this->options['format']]
        ));
    }

}
