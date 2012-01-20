# phpexcel

*PHPExcel module for Kohana 3.1.x*

- **Module URL:** <http://github.com/czukowski/kohana-phpexcel>
- **Compatible Kohana Version(s):** 3.1.x, 3.2.x

## Description

Kohana framework helper class to make spreadsheet creation easier

## Installation

Place the module in the modules/phpexcel.

The PHPExcel library (version 1.7.6) is now linked as a git submodule to vendor/phpexcel relative to the phpexcel module installation folder.

In the application/bootstrap.php add module loading
    
    Kohana::modules(array(
        ...
        'phpexcel'   => MODPATH.'phpexcel',
    ));

## Usage

Creating a Spreadsheet

    $ws = new Spreadsheet(array(
    	'author'       => 'Kohana-PHPExcel',
    	'title'	       => 'Report',
    	'subject'      => 'Subject',
    	'description'  => 'Description',
    ));
    
    $ws->set_active_sheet(0);
    $as = $ws->get_active_sheet();
    $as->setTitle('Report');
    
    $as->getDefaultStyle()->getFont()->setSize(9);
    
    $as->getColumnDimension('A')->setWidth(7);
    $as->getColumnDimension('B')->setWidth(40);
    $as->getColumnDimension('C')->setWidth(12);
    $as->getColumnDimension('D')->setWidth(10);
    
    $sh = array(
    	1 => array('Day','User','Count','Price'),
    	2 => array(1, 'John', 5, 587),
    	3 => array(2, 'Den', 3, 981),
    	4 => array(3, 'Anny', 1, 214)
    );
    
    $ws->set_data($sh, false);
    $ws->send(array('name'=>'report', 'format'=>'Excel5'));
