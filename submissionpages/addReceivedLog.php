<?php

$callToSave=strtoupper($_POST["receivedLog"]);
$response='';

//let's check to make sure it isn't already in the list before adding it again.
$currentList = file('logsRcvd.txt', FILE_IGNORE_NEW_LINES); //put existing list in array
if (!in_array($callToSave, $currentList)) { //the submitter's call isn't present
	file_put_contents('logsRcvd.txt', $callToSave.PHP_EOL , FILE_APPEND | LOCK_EX);
	$response="success";
} else {
	$response="fail";
}

echo $response;

?>