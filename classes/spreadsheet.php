<?php defined('SYSPATH') or die('No direct access allowed.');
/**
 * PHP Excel library. Helper class to make spreadsheet creation easier.
 *
 * @package    Spreadsheet
 * @author     Flynsarmy, Dmitry Shovchko
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 */
class Spreadsheet
{
	private $exts = array(
		'CSV'		=> 'csv',
		'PDF'		=> 'pdf',
		'Excel5' 	=> 'xls',
		'Excel2007' => 'xlsx',
	);
	private $mimes = array(
        'CSV' 		=> 'text/csv',
        'PDF' 		=> 'application/pdf',
        'Excel5' 	=> 'application/vnd.ms-excel',
        'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
	protected $options = array(
		'title'       => 'New Spreadsheet',
		'subject'     => 'New Spreadsheet',
		'description' => 'New Spreadsheet',
		'author'      => 'ClubSuntory',
		'format'      => 'Excel2007',
		'path'        => 'assets/downloads/spreadsheets/',
		'name'        => 'NewSpreadsheet',
	);
	/**
	 * @var PHPExcel
	 */
	protected $_spreadsheet;

	/**
	 * Creates the spreadsheet with given or default settings
	 * 
	 * @param array $options with optional parameters: title, subject, description, author
	 * @return void
	 */
	public function __construct($options = array())
	{
		/* PHP Excel integration */
		require_once Kohana::find_file('vendor', 'phpexcel/PHPExcel');

		$this->_spreadsheet = new PHPExcel();
		$this->set_options($options);
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
	 *  array(
	 *	   1 => array('A1', 'B1', 'C1', 'D1', 'E1'),
	 *	   2 => array('A2', 'B2', 'C2', 'D2', 'E2'),
	 *	   3 => array('A3', 'B3', 'C3', 'D3', 'E3'),
	 *  );
	 * 
	 * @param array of array( [row] => array([col]=>[value]) ) ie $arr[row][col] => value
	 * @return void
	 */
	public function set_data(array $data, $multi_sheet = FALSE)
	{
		// Single sheet ones can just dump everything to the current sheet
		if ( ! $multi_sheet)
		{
			$sheet = $this->_spreadsheet->getActiveSheet();
			$this->set_sheet_data($data, $sheet);
		}
		// Have to do a little more work with multi-sheet
		else
		{
			foreach ($data as $sheetname => $sheetData)
			{
				$sheet = $this->_spreadsheet->createSheet();
				$sheet->setTitle($sheetname);
				$this->set_sheet_data($sheetData, $sheet);
			}
			// Now remove the auto-created blank sheet at start of XLS
			$this->_spreadsheet->removeSheetByIndex(0);
		}
	}

	protected function set_options($options)
	{
		$this->options = Arr::merge($this->options, $options);
		return $this;
	}

	protected function set_properties()
	{
		$this->_spreadsheet->getProperties()
			->setCreator($this->options['author'])
			->setTitle($this->options['title'])
			->setSubject($this->options['subject'])
			->setDescription($this->options['description']);
		return $this;
	}

	protected function set_sheet_data(array $data, PHPExcel_Worksheet $sheet)
	{
		foreach ($data as $row => $columns)
		{
			foreach ($columns as $column => $value)
			{
				$sheet->setCellValueByColumnAndRow($column, $row, $value);
			}
		}
	}

	/**
	 * Writes spreadsheet to file
	 * 
	 * @param array $settings with optional parameters: format, path, name (no extension)
	 * @return Path to spreadsheet
	 */
	public function save($settings = array())
	{
		// Set document properties
		$this->set_properties();

		$settings = array_merge($this->options, $settings);

		// Generate full path
		$settings['fullpath'] = $settings['path'].$settings['name'].'_'.time().'.'.$this->exts[$settings['format']];

		$writer = PHPExcel_IOFactory::createWriter($this->_spreadsheet, $settings['format']);

		if ($settings['format'] == 'CSV')
		{
			$writer->setUseBOM(true);
		}
		$writer->save($settings['fullpath']);

		return $settings['fullpath'];
	}

	/**
	 * Send spreadsheet to browser
	 * 
	 * @param array $settings with optional parameters: format, name (no extension)
	 * @return void
	 */
	public function send($settings = array())
	{
		// Set document properties
		$this->set_properties();

		$settings = array_merge($this->options, $settings);

		$writer = PHPExcel_IOFactory::createWriter($this->_spreadsheet, $settings['format']);

		$ext = $this->exts[$settings['format']];
		$mime = $this->mimes[$settings['format']];

		$response = Request::current()->response();
		$response->headers(array(
			'Content-Type' => $mime,
			'Content-Disposition' => 'attachment;filename="'.$settings['name'].'.'.$ext.'"',
			'Cache-Control' => 'max-age=0',
		));
		$response->send_headers();

		if ($settings['format'] == 'CSV')
		{
			$writer->setUseBOM(TRUE);
		}

		$writer->save('php://output');
		exit;
	}

}