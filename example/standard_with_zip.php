<?php
require_once '../vendor/autoload.php';
$header = [
    ["Name","Username","Email", "Address","Phone","Company"]
];

$testerApi = "https://jsonplaceholder.typicode.com/users";
$client    = new \GuzzleHttp\Client();
$pathBase  = str_replace('example','',__DIR__).'storage';
$zip       = new ZipArchive();
$outputZipPath = $pathBase.'/standard_with_zip.zip';
$zip = new ZipArchive();
if ($zip->open($outputZipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== TRUE) {
    throw new \Exception("Cannot create zip file");
}
try {
    $request  = $client->get($testerApi);
    $response = json_decode($request->getBody()->getContents(),true);
    $xlsx = new \Anton\SimpleXlsx\SimpleXlsx($header,'standard_with_zip',null,1,null,$pathBase,null,null);
    $row  = $xlsx->setSpreadsheet();
    array_map(/**
     * @throws Exception
     */ function ($item) use(&$row,&$xlsx){
        $color = ($row%2) === 0;
        $color = false;
        $address = $item['address'];
        $addressComplete = $address['street'].' '.$address['suite'].','.$address['city'];
        $xlsx->setBodyCell(0,0,$row,$item['name'],$color);
        $xlsx->setBodyCell(0,1,$row,$item['username'],$color);
        $xlsx->setBodyCell(0,2,$row,$item['email'],$color);
        $xlsx->setBodyCell(0,3,$row,$addressComplete,$color);
        $xlsx->setBodyCell(0,4,$row,$item['phone'],$color);
        $xlsx->setBodyCell(0,5,$row,$item['company']['name'],$color);
        $row++;

    },$response);
    $xlsx->save($zip);
    $zip->close();
    if(file_exists($pathBase.'/STANDARD_WITH_ZIP.XLSX'))
        unlink($pathBase.'/STANDARD_WITH_ZIP.XLSX');

} catch (\GuzzleHttp\Exception\GuzzleException $e) {
    throw new \Exception($e->getMessage());
}


