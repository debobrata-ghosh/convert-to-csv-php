<?php 
  //Had to change this path to point to IOFactory.php.
  //Do not change the contents of the PHPExcel-1.8 folder at all.
  include('Classes/PHPExcel/IOFactory.php');

  //Use whatever path to an Excel file you need.
  $inputFileName = 'RS_Update_Reports_Template.xlsx';

  try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
  } catch (Exception $e) {
    die('Error loading file "' . pathinfo($inputFileName, PATHINFO_BASENAME) . '": ' . 
        $e->getMessage());
  }

  $sheet = $objPHPExcel->getSheet(0);
  $highestRow = $sheet->getHighestRow();
  $highestColumn = $sheet->getHighestColumn();
$result  = array();
$rowData12  = array();
  for ($row = 2; $row <= $highestRow; $row++) { 
  $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);
	$rowData12[] =	$rowData;									
  }	
$lic = array('Single','Multi-users','Enterprise');
 foreach($rowData12 as $row){
  for($i=0;$i<3;$i++){
		$j = $i+11;		
		$pdate = date("d-m-Y", strtotime("now"));
		$result[] = array( 
		'Previous Code' => $row[0][0],
		'New Previous Code' => $row[0][1],
		'Visible' => $row[0][2],
		'Findable' => $row[0][3],
		'On Demand' => $row[0][4],
		'File' => $row[0][5],
		'Sample File' => $row[0][6],
		'Image File' => $row[0][7],
		'Graph Image File' => $row[0][8],
		'Report Type' => $row[0][9],
		'Title' => $row[0][10],
		'Licence Type' => $lic[$i],
		'Price' => $row[0][$j],
		'Topic' => $row[0][14],
		'Sectors' => $row[0][15],
		'Hot topics' => $row[0][16],
		'Geography' => $row[0][17],
		'Number Of Pages' => $row[0][18],
		'Publication Date' => $pdate,
		'Synopsis' => $row[0][20],
		'Executive Summary' => $row[0][21],
		'Scope' => $row[0][22],
		'Reasons To Buy' => $row[0][23],
		'Key Highlights' => $row[0][24],
		'Keywords' => $row[0][25],
		'Companies Mentioned' => $row[0][26],
		'Table Of Contents' => $row[0][27],
		'List Of Tables' => $row[0][28],
		'List Of Figures' => $row[0][29],
		'Project Value' => $row[0][30],
		'project Stage' => $row[0][31],
		'Quote' => $row[0][32],
		'Quote Source' => $row[0][33],
		'Redirect URL' => $row[0][34],
		'Tags' => $row[0][35],
		'Topic_Id' => $row[0][36],
		'Manual URL' => $row[0][37],
		'Methodology' => $row[0][38]
		);		
		if($i==3) break;		
		}
	}   
    echo '<pre>';
			  print_r($result);
			echo '</pre>';  
	//Create a CSV file
	$file = fopen('Newfile.csv', 'w');
	foreach ($result as $line) {
		//put data into csv file
		fputcsv($file, $line);
	}
	fclose($file);		
			

?>