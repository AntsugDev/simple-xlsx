<?php
require_once '../vendor/autoload.php';
$header = [
    ["Name","Username","Email", "Address","Phone","Company"]
];

$testerApi = "https://jsonplaceholder.typicode.com/users";
$client    = new \GuzzleHttp\Client();
$pathBase  = str_replace('example','',__DIR__).'storage';
try {
    $request  = $client->get($testerApi);
    $response = json_decode($request->getBody()->getContents(),true);
    $xlsx = new \Anton\SimpleXlsx\SimpleXlsx($header,'fill',null,1,null,$pathBase,null,null);
    $row  = $xlsx->setSpreadsheet();
    array_map(/**
     * @throws Exception
     */ function ($item) use(&$row,&$xlsx){
        $color = ($row%2) === 0;
        $address = $item['address'];
        $fill = [
            "color"=> $color ? '48C9B0' : 'F5B041',
            "bold" => $color
        ];
        $addressComplete = $address['street'].' '.$address['suite'].','.$address['city'];
        $xlsx->setBodyCell(0,0,$row,$item['name'],$color,'left',null,$fill);
        $xlsx->setBodyCell(0,1,$row,$item['username'],$color,'left',null,$fill);
        $xlsx->setBodyCell(0,2,$row,$item['email'],$color,'center',null,$fill);
        $xlsx->setBodyCell(0,3,$row,$addressComplete,$color,'center',null,$fill);
        $xlsx->setBodyCell(0,4,$row,$item['phone'],$color,'right',null,$fill);
        $xlsx->setBodyCell(0,5,$row,$item['company']['name'],$color,'right',null,$fill);
        $row++;

    },$response);
    $xlsx->save();

} catch (\GuzzleHttp\Exception\GuzzleException $e) {
    throw new \Exception($e->getMessage());
}


