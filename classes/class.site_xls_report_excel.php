<?php

class Site_XLS_Report_Excel {

	protected $objPHPExcel;
	protected $o;
  protected $sheetIndex=0;

	const RGB_BACKGROUND_FIRST_ROW='e0eaf1';
	
	public function __construct($o){

		$this->o=$o;

		$this->objPHPExcel = new PHPExcel();
        $this->objPHPExcel->getProperties()->setCreator("Drupal");
        $this->objPHPExcel->setActiveSheetIndex($this->sheetIndex);
        $this->objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')->setSize(10);

    $this->add_modules();
    $this->add_content_types();
    $this->add_content_types_details();
    $this->add_roles_permission();
    $this->add_vocabularies();
    $this->add_taxonomies();
	}


  /**
   * Add the modules sheet (the first)
   *
   */
	public function add_modules(){

    	if(!isset($this->o->ml_info)){
    		return null;
    	}

      $range = "A1:H1";

    	$this->objPHPExcel->setActiveSheetIndex($this->sheetIndex);
    	$this->objPHPExcel->getActiveSheet()->setTitle('module_list');

		// Color
    $this->objPHPExcel->getActiveSheet()->getStyle( $range )->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => self::RGB_BACKGROUND_FIRST_ROW)
            )
        )
    );
    $this->objPHPExcel->getActiveSheet()->getStyle( $range  )->getFont()->setBold(true);


    // First row
    $i=1;
    $this->objPHPExcel->getActiveSheet()
        ->setCellValue('A'.$i, 'Name')
        ->setCellValue('B'.$i, 'Description')
        ->setCellValue('C'.$i, 'Core')
        ->setCellValue('D'.$i, 'Package')
        ->setCellValue('E'.$i, 'Version')
        ->setCellValue('F'.$i, 'Project')
        ->setCellValue('G'.$i, 'Dependencies')
        ->setCellValue('H'.$i, 'DateTime')
        ;

		$i=2;

		foreach($this->o->ml_info as $ml){

      $ml['dependencies'] = (isset($ml['dependencies'])) ? $ml['dependencies'] : '';
      $ml['package'] = (isset($ml['package'])) ? $ml['package'] : '';
			$ml['datestamp'] = (isset($ml['datestamp'])) ? date("Y-m-d", $ml['datestamp']) : '';

			$this->objPHPExcel->getActiveSheet()
				->setCellValue('A'.$i, $ml['name'])
        ->setCellValue('B'.$i, $ml['description'])
        ->setCellValue('C'.$i, $ml['core'])
        ->setCellValue('D'.$i, $ml['package'])
        ->setCellValue('E'.$i, $ml['version'])
        ->setCellValue('F'.$i, $ml['project'])
        ->setCellValue('G'.$i, $ml['dependencies'])
        ->setCellValue('H'.$i, $ml['datestamp'] )
        ;

      $this->objPHPExcel->getActiveSheet()->getStyle('H'.$i)->
        getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD2);

      $i++;

		}

    foreach(range('A','H') as $L){
        $this->objPHPExcel->getActiveSheet()->getColumnDimension( $L )->setAutoSize(true);
    }

    $this->sheetIndex++;

	}


  public function add_content_types(){

    if(!isset($this->o->cts)){
        return null;
      }

    $sheetId = $this->sheetIndex;
    $this->objPHPExcel->createSheet(NULL, $sheetId);
    $this->objPHPExcel->setActiveSheetIndex($sheetId);
    $this->objPHPExcel->getActiveSheet()->setTitle('content_types');

    // Color
    $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => self::RGB_BACKGROUND_FIRST_ROW)
            )
        )
    );
    $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->getFont()->setBold(true);


    // First row

    $i=1;

    $this->objPHPExcel->getActiveSheet()
        ->setCellValue('A'.$i, 'Name')
        ->setCellValue('B'.$i, 'Machine name')
        ->setCellValue('C'.$i, 'Base')
        ->setCellValue('D'.$i, 'Module')
        ->setCellValue('E'.$i, 'Description')
    ;

    $i=2;

    foreach($this->o->cts as $ct){

      $this->objPHPExcel->getActiveSheet()
        ->setCellValue('A'.$i, $ct->name)
        ->setCellValue('B'.$i, $ct->type)
        ->setCellValue('C'.$i, $ct->base)
        ->setCellValue('D'.$i, $ct->module)
        ->setCellValue('E'.$i, strip_tags($ct->description))
        ;

      $i++;
    }

    foreach(range('A','E') as $L){
        $this->objPHPExcel->getActiveSheet()->getColumnDimension( $L )->setAutoSize(true);
    }

    $this->sheetIndex++;

  }


  public function add_roles_permission(){

    if(!isset($this->o->permissions) || !isset($this->o->roles)){
      return null;
    }

    // Determine the admin
    $kadmin = array_search('administrator', $this->o->roles);

    $list_all = array_keys($this->o->permissions[$kadmin]);

    $sheetId = $this->sheetIndex;
    $this->objPHPExcel->createSheet(NULL, $sheetId);
    $this->objPHPExcel->setActiveSheetIndex($sheetId);
    $this->objPHPExcel->getActiveSheet()->setTitle('permissions');


    // First row

    $i=1;

    $aa = range('A','Z');


    $this->objPHPExcel->getActiveSheet()->setCellValue('A' . $i, 'Permission');
    
    $rc = 0;
    foreach($this->o->roles as $R){
      $this->objPHPExcel->getActiveSheet()->setCellValue($aa[ $rc+1 ] . $i, $R);
      $rc++;
    }

    $max = $aa[ count($this->o->roles)];
    $range = "A1:".$max."1";

    // Color
    $this->objPHPExcel->getActiveSheet()->getStyle( $range )->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => self::RGB_BACKGROUND_FIRST_ROW)
            )
        )
    );
    $this->objPHPExcel->getActiveSheet()->getStyle( $range )->getFont()->setBold(true);



    $i=2;

    for($x=0; $x<count($list_all); $x++){

      $this->objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $list_all[$x]);

      for($j=0;$j<=count($this->o->permissions);$j++){


        if(isset($this->o->permissions[$j][ $list_all[$x] ]) && $this->o->permissions[$j][ $list_all[$x] ] == TRUE){
          $this->objPHPExcel->getActiveSheet()->setCellValue($aa[ $j ] . $i, 'X');
        }
        else{
          // $this->objPHPExcel->getActiveSheet()->setCellValue($aa[ $j ] . $i, '');
        }
      }

      $i++;
    }

    foreach(range('A',$max) as $L){
        $this->objPHPExcel->getActiveSheet()->getColumnDimension( $L )->setAutoSize(true);
    }

    $this->sheetIndex++;

  }


  public function add_content_types_details(){

    if(!isset($this->o->ct_fields)){
      return null;
    }

    foreach($this->o->ct_fields as $ctf){


      $bo = current($ctf);
      $bundle_name = $bo['bundle'];

      // create sheet
      $sheetId = $this->sheetIndex;
      $this->objPHPExcel->createSheet(NULL, $sheetId);
      $this->objPHPExcel->setActiveSheetIndex($sheetId);
      $this->objPHPExcel->getActiveSheet()->setTitle('content_types_'.$bundle_name);

      // Color
      $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->applyFromArray(
          array(
              'fill' => array(
                  'type' => PHPExcel_Style_Fill::FILL_SOLID,
                  'color' => array('rgb' => self::RGB_BACKGROUND_FIRST_ROW)
              )
          )
      );
      $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->getFont()->setBold(true);

      // First row
      
      $i=1;

      $this->objPHPExcel->getActiveSheet()
          ->setCellValue('A'.$i, 'Label')
          ->setCellValue('B'.$i, 'Field name')
          ->setCellValue('C'.$i, 'Widget type')
          ->setCellValue('D'.$i, 'Required')
          ->setCellValue('E'.$i, 'Default value')
      ;


      $i=2;

      foreach($ctf as $ct){

        $ct['default_value'] = (isset($ct['default_value'])) ? $ct['default_value'] : '';

        $this->objPHPExcel->getActiveSheet()
          ->setCellValue('A'.$i, $ct['label'])
          ->setCellValue('B'.$i, $ct['field_name'])
          ->setCellValue('C'.$i, $ct['widget']['type'])
          ->setCellValue('D'.$i, $ct['required'])
          ->setCellValue('E'.$i, $ct['default_value'])
          ;

        $i++;
      }
      

      $this->sheetIndex++;

      foreach(range('A','E') as $L){
          $this->objPHPExcel->getActiveSheet()->getColumnDimension( $L )->setAutoSize(true);
      }
    }
  }


  public function add_taxonomies(){

    if(!isset($this->o->tax)){
      return null;
    }

    $taxs = $this->o->tax;

    foreach($this->o->vocabularies as $v){

      $sheetId = $this->sheetIndex;
      $this->objPHPExcel->createSheet(NULL, $sheetId);
      $this->objPHPExcel->setActiveSheetIndex($sheetId);
      $this->objPHPExcel->getActiveSheet()->setTitle('tax_'.$v->machine_name);

      // Color
      $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->applyFromArray(
          array(
              'fill' => array(
                  'type' => PHPExcel_Style_Fill::FILL_SOLID,
                  'color' => array('rgb' => self::RGB_BACKGROUND_FIRST_ROW)
              )
          )
      );
      $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->getFont()->setBold(true);

      // First row

      $i=1;

      $this->objPHPExcel->getActiveSheet()
          ->setCellValue('A'.$i, 'Tid')
          ->setCellValue('B'.$i, 'Vid')
          ->setCellValue('C'.$i, 'Name')
          ->setCellValue('D'.$i, 'Description')
          ->setCellValue('E'.$i, 'uuid')
      ;

      $i=2;

      $taxs = taxonomy_get_tree($v->vid);

      foreach($taxs as $tax){

        $uuid = (isset($tax->uuid)) ? $tax->uuid : '';

        $this->objPHPExcel->getActiveSheet()
          ->setCellValue('A'.$i, $tax->tid)
          ->setCellValue('B'.$i, $tax->vid)
          ->setCellValue('C'.$i, $tax->name)
          ->setCellValue('D'.$i, strip_tags($tax->description))
          ->setCellValue('E'.$i, $uuid)
          ;

        $i++;
      }

      foreach(range('A','E') as $L){
          $this->objPHPExcel->getActiveSheet()->getColumnDimension( $L )->setAutoSize(true);
      }

      $this->sheetIndex++;
    }

  }


  public function add_vocabularies(){

    if(!isset($this->o->vocabularies)){
        return null;
      }

    $sheetId = $this->sheetIndex;
    $this->objPHPExcel->createSheet(NULL, $sheetId);
    $this->objPHPExcel->setActiveSheetIndex($sheetId);
    $this->objPHPExcel->getActiveSheet()->setTitle('vocabularies');

    // Color
    $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->applyFromArray(
        array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => self::RGB_BACKGROUND_FIRST_ROW)
            )
        )
    );
    $this->objPHPExcel->getActiveSheet()->getStyle( "A1:E1" )->getFont()->setBold(true);


    // First row

    $i=1;

    $this->objPHPExcel->getActiveSheet()
        ->setCellValue('A'.$i, 'Vid')
        ->setCellValue('B'.$i, 'Machine name')
        ->setCellValue('C'.$i, 'Name')
        ->setCellValue('D'.$i, 'Description')
        ->setCellValue('E'.$i, 'Parent')
    ;

    $i=2;

    foreach($this->o->vocabularies as $v){

      $this->objPHPExcel->getActiveSheet()
        ->setCellValue('A'.$i, $v->vid)
        ->setCellValue('B'.$i, $v->machine_name)
        ->setCellValue('C'.$i, $v->name)
        ->setCellValue('D'.$i, strip_tags($v->description))
        ->setCellValue('E'.$i, $v->hierarchy)
        ;

      $i++;
    }

    foreach(range('A','E') as $L){
        $this->objPHPExcel->getActiveSheet()->getColumnDimension( $L )->setAutoSize(true);
    }

    $this->sheetIndex++;

  }

	public function stream($filename){
        
      // Imposta il primo come foglio attivo
      $this->objPHPExcel->setActiveSheetIndex(0);
      
      // stream file
      ECC_ExcelWriter::send(
      									$this->objPHPExcel, 
                        'Excel5', 
                        $filename
                      );
	}


}


class ECC_ExcelWriter extends PHPExcel_IOFactory {
    
    
    public function send($obj, $type, $filename){
        
      date_default_timezone_set('Europe/Rome');
      
      switch($type){
          
        case 'Excel2007':

            // Redirect output to a client’s web browser (Excel2007)
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
            header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header ('Pragma: public'); // HTTP/1.0

            $objWriter = self::createWriter($obj, $type);
            $objWriter->save('php://output');
        
        break;
    
    
        case 'Excel5':
        default: 
            
            // Redirect output to a client’s web browser (Excel5)
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
            header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header ('Pragma: public'); // HTTP/1.0

            $objWriter = self::createWriter($obj, 'Excel5');
            $objWriter->save('php://output');
            
        break;
      
      }
    }
}
