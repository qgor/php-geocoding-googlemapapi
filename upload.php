<?php
ini_set('max_execution_time', 1800);
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
//define('Geocode_API_Key', 'API-KEY-GOES-HERE'); 
$error = array();

function writeLog($log){
	echo date('H:i:s')." ".$log.EOL;
}

// Encode a string to URL-safe base64
function encodeBase64UrlSafe($value)
{
  return str_replace(array('+', '/'), array('-', '_'),
    base64_encode($value));
}

// Decode a string from URL-safe base64
function decodeBase64UrlSafe($value)
{
  return base64_decode(str_replace(array('-', '_'), array('+', '/'),
    $value));
}

// Sign a URL with a given crypto key
// Note that this URL must be properly URL-encoded
function signUrl($myUrlToSign, $privateKey)
{
  // parse the url
  $url = parse_url($myUrlToSign);

  $urlPartToSign = $url['path'] . "?" . $url['query'];

  // Decode the private key into its binary format
  $decodedKey = decodeBase64UrlSafe($privateKey);

  // Create a signature using the private key and the URL-encoded
  // string using HMAC SHA1. This signature will be binary.
  $signature = hash_hmac("sha1",$urlPartToSign, $decodedKey,  true);

  $encodedSignature = encodeBase64UrlSafe($signature);

  return $myUrlToSign."&sensor=false&key=".$encodedSignature;
}

function createSQL($store, $raw, $storeType, $file = null){
	$url = "https://maps.googleapis.com/maps/api/geocode/json?address=";
	$url .= urlencode($store['addr']).",".urlencode(" ".$store['city']);
	if ($store['region'] != "")
		$url .= ",".urlencode(" ".$store['region']);
	$url .= ",".urlencode(" ".$store['country']);
	$url .= "&sensor=false";
	$url .= "&key=".Geocode_API_Key;

	$store["lat"] = "";
    $store["lng"] = "";

	//$url = signUrl($url, Geocode_API_Key);
//	echo $url.EOL;
/*	
	$ch = curl_init($url);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$result = curl_exec($ch);
	$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
	curl_close($ch);
	echo "-".$result."-".EOL;
	$json = json_decode($result);
*/
/**/
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_URL, $url);
	curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$response = curl_exec($ch);
//	echo "-".$response."-".EOL;
	$json = json_decode($response, true);

	if ($json['status'] != 'OK') {
		$store["error"] = "Geocoding failed - ".$json['status'];
		array_push($GLOBALS["error"], $raw);
		return;
	}

//	print_r($json);

	$geometry = $json['results'][0]['geometry']; 
	$store["lat"] = $geometry['location']['lat'];
    $store["lng"] = $geometry['location']['lng'];
/**/

	/*
  `store_type` varchar(255) NOT NULL,
  `name` varchar(255) NOT NULL,
  `address` varchar(255) NOT NULL,
  `city` varchar(255) NOT NULL,
  `province` varchar(255) NOT NULL,
  `postalcode` varchar(255) NOT NULL,
  `phone` varchar(255) NOT NULL,
  `info` longtext NOT NULL,
  `lat` varchar(255) NOT NULL,
  `long` varchar(255) NOT NULL,
	*/
	
	$sql_format = "INSERT INTO storelocator (`store_type`, `name`, `address`, `city`, `province`, `country`, `postalcode`, `phone`, `info`, `lat`, `long`) VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '', '%s', '%s');\n";
	$sql = sprintf($sql_format, $storeType, $store["name"], $store["addr"], $store["city"], $store["region"], $store["country"], $store["postal"], $store["phone"], $store["lat"], $store["lng"]);
	if ($file){
		fwrite($file, $sql);
	}
	else{
		echo $sql;
	}
}

require_once 'Classes/PHPExcel/IOFactory.php';

PHPExcel_Settings::setZipClass(PHPExcel_Settings::PCLZIP);

$objPHPExcel = PHPExcel_IOFactory::load("list.xlsx");

writeLog( "Loading from xlsx file..." );
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load("list.xlsx");
writeLog( "Loaded." );

$link = mysql_connect('localhost', 'db-user', 'db-password')
    OR die(mysql_error());

$start = 2001;
$end = 4000;
	
$file = fopen("output.sql", "w");

foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
	
	if ($worksheet->getTitle() != "Retail Partners")
		continue;

	writeLog( "Retail Partners sheet found." );

	foreach ($worksheet->getRowIterator() as $row) {
	
		if ($row->getRowIndex() < $start)
			continue;

		if ($row->getRowIndex() > $end)
			break;
			
		$cellIterator = $row->getCellIterator();
		$cellIterator->setIterateOnlyExistingCells(false); // Loop all cells, even if it is not set

		$colIndex = 0;
		$store = array();
		$raw = array();
		$store['row'] = $row->getRowIndex();

		foreach ($cellIterator as $cell) {
		
			$val = "";
			if (!is_null($cell)) 
				$val = $cell->getCalculatedValue();
			$val = mysql_real_escape_string($val);
			
			switch($colIndex){
				case 0:
					$store['name'] = $val;
					break;
				case 1:
					$store['addr'] = $val;
					break;
				case 2:
					$store['city'] = $val;
					break;
				case 3:
					if (strtolower($val) == "other")
						$store['region'] = "";
					else
						$store['region'] = $val;
					break;
				case 4:
					$store['country'] = $val;
					break;
				case 5:
					$store['postal'] = $val;
					break;
				case 6:
					$store['phone'] = $val;
					break;
			}

			$raw[$colIndex] = $val;
			$colIndex++;
			
		}

		createSQL($store, $raw, "2", $file);
//		sleep(3);
	}
}

fclose($file);

$fileError = fopen('error.csv', 'w');

foreach ($error as $item) {
    fputcsv($fileError, $item);
}

fclose($fileError);

writeLog( "Scanning completed." );

?>


