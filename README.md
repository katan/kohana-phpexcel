# phpexcel

*PHPExcel module for Kohana 3.2.x*

- **Module URL:** <http://github.com/rafsoaken/kohana-phpexcel>
- **Compatible Kohana Version(s):** 3.2.x (not tested in 3.1.x)

## Description

Kohana framework helper class to make and read spreadsheet easier
Added in this fork:
- read spreadsheet from csv, excel5 and excel2007

## Installation



*Via git submodules (recommended):*
The PHPExcel library (version 1.7.6) is now linked as a git submodule to vendor/phpexcel relative to the phpexcel module installation folder.
Clone the repository via:

    git submodule add git://github.com/yuankai/kohana-phpexcel.git modules/phpexcel
    git submodule init

Then update the git submodule and the contained vendor/phpexcel sub-submodule from the repository root:

    git submodule update --init --recursive

If you followed the commands exactly, all contents should have been cloned by now. Further information for git submodules within submodules on [Stackoverflow] [1].
Finally load the module in your application (see below).

*Via ZIP file download:*
Download and extract the zip file. Place the phpexcel module in modules/phpexcel.
Because the git submodule in the vendor/phpexcel folder is not included in the download, please go and get it at
<https://github.com/rafsoaken/phpexcel> (again, the zip download is what you want), then replace the empty vendor/phpexcel folder with it.

Finally load the module in your application as follows:

In the application/bootstrap.php add module loading
    
    Kohana::modules(array(
        ...
        'phpexcel'   => MODPATH.'phpexcel',
    ));

## Usage

Creating a Spreadsheet

    $spreadsheet = Spreadsheet::factory(array(
          'author'  => 'Kohana-PHPExcel',
          'title'      => 'Report',
          'subject' => 'Subject',
          'description'  => 'Description',
          'path' => '/',
          'name' => 'report'
    ));
    $spreadsheet->set_active_sheet(0);
    $as = $spreadsheet->get_active_sheet();
    $as->setTitle('Consumos');
    $as->getDefaultStyle()->getFont()->setSize(9);

    $as->getStyle('A1:G1')->applyFromArray(Kohana::$config->load('styles.header'));
    $as->getRowDimension(1)->setRowHeight(24);
    $as->getStyle('A1:G1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $as->getStyle('A1:G1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    $as->getColumnDimension('A')->setWidth(8);
    $as->getColumnDimension('B')->setWidth(12);
    $as->getColumnDimension('C')->setWidth(46);
    $as->getColumnDimension('D')->setWidth(36);
    
    $sh = array(
    	1 => array('Day','User','Count','Price'),
    	2 => array(1, 'John', 5, 587),
    	3 => array(2, 'Den', 3, 981),
    	4 => array(3, 'Anny', 1, 214)
    );
    
    $spreadsheet->set_data($sh, false);
    $spreadsheet->send();

Reading a Spreadsheet

    $spreadsheet = Spreadsheet::factory(
              array(
                        'filename' => 'spreadsheet.xlsx'
              ), FALSE)
              ->load()
              ->read();
    foreach ($spreadsheet as $v)
    {
              echo $v['A'].',';
    }

[1]: http://stackoverflow.com/questions/1535524/git-submodule-inside-of-a-submodule        "Stackoverflow"
