<?php
namespace App;

class PHPExcelSimplification{
	static private $PHPExcelInstance = null;
	static private $autoWidthCell = false;

	static public function generateExcel($data){
		$response = new \stdClass();

		try{
			self::validateConfigData($data);
			self::initPHPExcelInstace();
			self::render($data);

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

	static private function resetClassData(){
		self::$PHPExcelInstance = null;
		self::$autoWidthCell = false;
	}

	static private function render($data){
		self::makeStructure($data);
		self::makeFile($data['url']);
	}

	static private function makeFile($url){
		(\PHPExcel_IOFactory::createWriter(self::$PHPExcelInstance, 'Excel2007'))->save($url);
		if(!file_exists($url)) throw new \Exception('Error al momento de crear el archivo.');
	}

	static private function makeStructure($data){
		if(isset($data['config'])) self::setOptions($data['config']);
		self::setCellsValues($data['values']);
	}

	static function setCellsValues($cells){
		foreach($cells as $cell => $content){
			self::$PHPExcelInstance->getActiveSheet()->setCellValue($cell,$content);
			if(self::$autoWidthCell){
				self::$PHPExcelInstance->getActiveSheet()->getColumnDimension($cell)->setAutoSize(true);
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
						self::$PHPExcelInstance->getActiveSheet()->mergeCells($cell);
					break;
					case 'background':
						self::$PHPExcelInstance->getActiveSheet()->getStyle($cell)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB($data[1]);	
					break;
					case 'color':
						self::$PHPExcelInstance->getActiveSheet()->getStyle($cell)->getFont()->getColor()->setRGB($data[1]);
					break;
					case 'alignText':
						if($data[1] == 'center'){
							self::$PHPExcelInstance->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
						}else{
							throw new \Exception('Configuración de alineación invalida.');
						}
					break;
					case 'bold':
						self::$PHPExcelInstance->getActiveSheet()->getStyle($cell)->getFont()->setBold(true);
					break;
					case 'border':
						if($data[1] == 'medium'){
							self::$PHPExcelInstance->getActiveSheet()->getStyle($cell)->getBorders()->applyFromArray([
								'bottom' => [
									'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								],
								'top' => [
									'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								],
								'left' => [
									'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								],
								'right' => [
									'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								]
							]);
						}else{
							throw new \Exception('Configuración de bordes invalida.');
						}
					break;
					case 'borderBottom' :
						if($data[1] == 'thin'){
							self::$PHPExcelInstance->getActiveSheet()->getStyle($cell)->getBorders()->applyFromArray([
								'bottom' => [
									'style' => \PHPExcel_Style_Border::BORDER_THIN,
									'color' => [
										'rgb' => $data[2] ?? '000000'
									]
								]
							]);
						}else{
							throw new \Exception('Configuración de bordes invalida.');
						}
					break;
					case 'widthColumn':
						self::$PHPExcelInstance->getActiveSheet()->getColumnDimension(preg_replace('/[0-9]+/', '',$cell))->setWidth($data[1]); 
					break;
					default:
						throw new \Exception('Opción de configuración de celda no reconocida.');
					break;
				}
			}
		}
	}

	static private function setGlobalOptions($configs){
		foreach($configs as $config){
			switch($config){
				case 'noGridLines':
					self::$PHPExcelInstance->getActiveSheet()->setShowGridlines(false);
				break;
				case 'autoWidthCell':
					self::$autoWidthCell = true;
				break;
				default:
					throw new \Exception('Opción global de configuración no reconocida.');
				break;
			}
		}
	}

	static private function validateConfigData($config){
		if(!isset($config['values']) || !isset($config['url'])) throw new \Exception('Data de configuración invalida');
	}

	static private function initPHPExcelInstace(){
		if(!class_exists('PHPExcel')) require '../vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
		self::$PHPExcelInstance = new \PHPExcel();
		self::setAttributesExcel();
		self::$PHPExcelInstance->setActiveSheetIndex(0);
	}

	static private function setAttributesExcel(){
		self::$PHPExcelInstance->getProperties()
		->setCreator("Suplos")
		->setLastModifiedBy("Suplos");
	}
}