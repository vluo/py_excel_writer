<?php
	
$start = microtime(1);	
$file = 'E:\vluos\docker\python3\httpserver\excels\data.json';
$url = 'http://127.0.0.1:88/write_excel';
ini_set('memory_limit', '500M');
$data_string = file_get_contents($file);
//die($data_string);
$ch = curl_init();

curl_setopt($ch, CURLOPT_URL, $url);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);
curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, FALSE);

// POST数据
curl_setopt($ch, CURLOPT_POST, 1);
// 把post的变量加上
curl_setopt($ch,  CURLOPT_POSTFIELDS, $data_string);
curl_setopt($ch, CURLOPT_HTTPHEADER, array(
    'Content-Type: application/json',
    'Content-Length: ' . strlen($data_string))
);


$output = curl_exec($ch);

curl_close($ch);
echo "time tooke =".(microtime(1)-$start);
var_dump($output);