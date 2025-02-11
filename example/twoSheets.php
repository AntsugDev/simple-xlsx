<?php
require_once '../vendor/autoload.php';
$sheets = ["User","Company and Address"];
$header = [
    ["Name","Username","Email","Phone"],
    ["Name","Address","Company"],
];

$testerApi = "https://jsonplaceholder.typicode.com/users";
$client    = new \GuzzleHttp\Client();
$pathBase  = str_replace('example','',__DIR__).'storage';
try {
    $request  = $client->get($testerApi);
    $response = json_decode($request->getBody()->getContents(),true);
    $xlsx = new \Anton\SimpleXlsx\SimpleXlsx($header,'twoSheets',$sheets,1,null,$pathBase,null,null);
    $row  = $row1 =  $xlsx->setSpreadsheet();
    array_map(/**
     * @throws Exception
     */ function ($item) use(&$row,&$row1,&$xlsx){
        $color = ($row%2) === 0;
        $colorBis = ($row1%2) === 0;
        $address = $item['address'];
        $addressComplete = $address['street'].' '.$address['suite'].','.$address['city'];
        $xlsx->setBodyCell(0,0,$row,$item['name'],$color);
        $xlsx->setBodyCell(0,1,$row,$item['username'],$color);
        $xlsx->setBodyCell(0,2,$row,$item['email'],$color);
        $xlsx->setBodyCell(0,3,$row,$item['phone'],$color);
        $row++;

        $xlsx->setBodyCell(1,0,$row1,$item['name'],$colorBis);
        $xlsx->setBodyCell(1,1,$row1,$addressComplete,$colorBis);
        $xlsx->setBodyCell(1,2,$row1,$item['company']['name'],$colorBis);
        $row1++;
    },$response);
    $xlsx->save();

} catch (\GuzzleHttp\Exception\GuzzleException $e) {
    throw new \Exception($e->getMessage());
}


