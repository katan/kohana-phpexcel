# phpexcel

*PHPExcel module for Kohana 3.1.x, 3.2.x*

- **Module URL:** <http://github.com/rafsoaken/kohana-phpexcel>
- **Compatible Kohana Version(s):** 3.1.x, 3.2.x

## Description

Kohana framework helper class to make spreadsheet creation easier

## Installation



*Via git submodules (recommended):*
The PHPExcel library (version 1.7.6) is now linked as a git submodule to vendor/phpexcel relative to the phpexcel module installation folder.
Clone the repository via:

    git submodule add git://github.com/rafsoaken/kohana-phpexcel.git modules/phpexcel
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


[1]: http://stackoverflow.com/questions/1535524/git-submodule-inside-of-a-submodule        "Stackoverflow"
