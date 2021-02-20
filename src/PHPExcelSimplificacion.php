<?php

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class PHPSpreadsheetSimplification{
	static private $SpreadsheetInstance = null;
	static private $SpreadsheetDrawingInstance = null;
	static private $autoWidthCell = false;

	static public function generateExcel($data){
		$response = new \stdClass();
		try{
			self::process($data);
			$response->bool = true;
			$response->msg = 'Formato generado satifactoriamente en la ruta especificada.';
		}catch(\Exception $e){
			$response->bool = false;
			$response->msg = 'No fue posible generar el formato.';
			$response->devMsg = $e->getMessage();
		}

		self::resetClassData();
		return $response;
	}

	static private function process($data){
		if(self::arrayIsAssoc($data)){
			self::validateConfigData($data);
			self::initPHPExcelInstace($data);
			self::render($data);
		}else{
			$indexPage = 0;
			foreach($data as $configPage){
				self::validateConfigData($configPage);

				if($indexPage <= 0) self::initPHPExcelInstace($configPage);
				else self::createNewPage($configPage,$indexPage);
				
				self::render($configPage);
				$indexPage++;
			}
		}
	}

	static private function createNewPage($configPage,$indexPage){
		self::$SpreadsheetInstance->createSheet();
		self::$SpreadsheetInstance->setActiveSheetIndex($indexPage);

		if(isset($configPage['config'])){
			if(isset($configPage['config']['titlePage'])) self::setTitlePage($configPage['config']['titlePage']);
		}
	}

	static private function resetClassData(){
		self::$SpreadsheetInstance = null;
		self::$SpreadsheetDrawingInstance = null;
		self::$autoWidthCell = false;
	}

	static private function render($data){
		self::makeStructure($data);
		self::makeFile($data['url']);
	}

	static private function makeFile($url){
		(new Xlsx(self::$SpreadsheetInstance, 'Excel2007'))->save($url);
		if(!file_exists($url)) throw new \Exception('Error al momento de crear el archivo.');
	}

	static private function makeStructure($data){
		if(isset($data['config'])) self::setOptions($data['config']);
		self::setCellsValues($data['values']);
	}

	static private function setCellsValues($cells){
		foreach($cells as $cell => $content){
			$data = explode('|',$content);
			if(isset($data[1])){
				if($data[1] == 'image'){
					self::renderImage($data,$cell);
				}else{
					throw new \Exception('Configuración de valor invalido.');
				}
			}else{
				self::$SpreadsheetInstance->getActiveSheet()->setCellValue($cell,$data[0]);
			}
			if(self::$autoWidthCell){
				self::$SpreadsheetInstance->getActiveSheet()->getColumnDimension($cell)->setAutoSize(true);
			}
		}
	}

	static private function renderImage($data,$cell){
		self::$SpreadsheetDrawingInstance = new Drawing();
		self::$SpreadsheetDrawingInstance->setPath($data[0]);
		self::$SpreadsheetDrawingInstance->setCoordinates($cell);
		
		if(count($data) >= 2){
			self::renderOptionsImage(array_slice($data,2));
		}

		self::$SpreadsheetDrawingInstance->setWorksheet(self::$SpreadsheetInstance->getActiveSheet());
	}

	static private function renderOptionsImage($options){
		foreach($options as $option){
			$data = explode(':',$option);
			if(count($data) < 2) throw new \Exception('Configuración de imagen invalida.');

			switch($data[0]){
				case 'heigth':
					self::$SpreadsheetDrawingInstance->setHeight($data[1]);
				break;
				case 'width':
					self::$SpreadsheetDrawingInstance->setWidth($data[1]);
				break;
				default:
					throw new \Exception('Configuración de imagen no reconocida.');
				break;
			}
		}
	}

	static private function setOptions($configs){
		if(isset($configs['global'])) self::setGlobalOptions($configs['global']);
		if(isset($configs['cells'])) self::setCellsOptions($configs['cells']);
	}

	static private function setCellsOptions($configs){
		foreach($configs as $cell => $config){
			$arrayConfigs = explode('|',$config);

			foreach($arrayConfigs as $singleConfig){
				$data = explode(':',$singleConfig);
				switch($data[0]){
					case 'merge':
						self::$SpreadsheetInstance->getActiveSheet()->mergeCells($cell);
					break;
					case 'background':
						self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB($data[1]);	
					break;
					case 'color':
						self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getFont()->getColor()->setRGB($data[1]);
					break;
					case 'alignText':
						if($data[1] == 'center'){
							self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
						}else if($data[1] == 'right'){
							self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
						}else if($data[1] == 'left'){
							self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
						}else{
							throw new \Exception('Configuración de alineación invalida ('.$data[1].').');
						}
					break;
					case 'bold':
						self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getFont()->setBold(true);
					break;
					case 'italic':
						self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getFont()->setItalic(true);
					break;
					case 'border':
						switch($data[1]){
							case 'medium':
								self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->applyFromArray([
									'borderStyle' => Border::BORDER_MEDIUM,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								]);
							break;
							case 'thin':
								self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->applyFromArray([
									'borderStyle' => Border::BORDER_THIN,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								]);
							break;
							default:
								throw new \Exception('Configuración de bordes invalida ('.$data[1].').');
							break;
						}
					break;
					case 'borderBottom' :
						if($data[1] == 'thin'){
							self::$SpreadsheetInstance->getActiveSheet()->getStyle($cell)->getBorders()->getBottom()->applyFromArray([
								'borderStyle' => Border::BORDER_THIN,
								'color' => [
									'rgb' => $data[2] ?? '000000'
								]
							]);
						}else{
							throw new \Exception('Configuración de bordes invalida ('.$data[1].').');
						}
					break;
					case 'widthColumn':
						self::$SpreadsheetInstance->getActiveSheet()->getColumnDimension(preg_replace('/[0-9]+/', '',$cell))->setWidth($data[1]); 
					break;
					default:
						throw new \Exception('Opción de configuración de celda no reconocida ('.$data[0].').');
					break;
				}
			}
		}
	}

	static private function setGlobalOptions($configs){
		foreach($configs as $config){
			switch($config){
				case 'noGridLines':
					self::$SpreadsheetInstance->getActiveSheet()->setShowGridlines(false);
				break;
				case 'autoWidthCell':
					self::$autoWidthCell = true;
				break;
				default:
					throw new \Exception('Opción global de configuración no reconocida ('.$config.').');
				break;
			}
		}
	}

	static private function validateConfigData($config){
		if(!isset($config['values']) || !isset($config['url'])) throw new \Exception('Data de configuración invalida son requeridos los indices values y url');
	}

	static private function initPHPExcelInstace($config){
		self::$SpreadsheetInstance = new Spreadsheet();
		self::setAttributesExcel();
		self::$SpreadsheetInstance->setActiveSheetIndex(0);

		if(isset($config['config'])){
			if(isset($config['config']['titlePage'])) self::setTitlePage($config['config']['titlePage']);
		}
		
	}

	static private function setAttributesExcel(){
		self::$SpreadsheetInstance->getProperties()
		->setCreator("")
		->setLastModifiedBy("");
	}

	static private function setTitlePage($title){
		self::$SpreadsheetInstance->getActiveSheet()->setTitle($title);
	}

	static private function arrayIsAssoc($arr){
		return array_keys($arr) !== range(0, count($arr) - 1);
	}
}
