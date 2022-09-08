<?php

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

require 'PHPMailer/src/Exception.php';
require 'PHPMailer/src/PHPMailer.php';
require 'PHPMailer/src/SMTP.php';
require_once 'PHPExcel/Classes/PHPExcel.php';


//First things first, let's do some security around the file upload
//check for the existence of a file, if empty, error and exit
if(!isset($_FILES["fileLog"])) {
	$fatalError = '';
	$fatalError .= 'There was a problem uploading your file. <br /><br />Please <a href="logupload.htm">try again</a>.';
	exit($fatalError);
}


// check to see if they have saved it in newer Office version resulting in .xlsx type.
if (in_array(mime_content_type($_FILES['fileLog']['tmp_name']),
	 array(
		'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	 )
	)) {
	// echo mime_content_type($_FILES['fileLog']['tmp_name']);
	$fatalError = '';
	$fatalError .= 'Improper file type uploaded. You can only upload a Microsoft Excel (.xls) file.';
	$fatalError .= '<br /><br />Please use the Excel Save As function to resave as "Excel 97-2003 Workbook (*.xls)" and upload again.';
	$fatalError .= '<br /><br />If the problem persists, please contact the DX Marathon administrator (k9el at dxmarathon dot com)';
	$fatalError .= '<br /><br /><a href="logupload.htm">Try again</a>';
    exit($fatalError);
}
//check the type of file being uploaded, anything other than Excel, should be trapped, error and exit
if (!in_array(mime_content_type($_FILES['fileLog']['tmp_name']), 
	 array(
		'application/vnd.ms-excel',
		'application/vnd.ms-office'
	 )
	)) {
	// echo mime_content_type($_FILES['fileLog']['tmp_name']);
	$fatalError = '';
	$fatalError .= 'Improper file type uploaded. You can only upload a Microsoft Excel (.xls) file.';
	$fatalError .= '<br /><br />Please consult the <a href="submission.htm">submission guidelines</a>.';
	$fatalError .= '<br /><br />If the problem persists, please contact the DX Marathon administrator (k9el at dxmarathon dot com)';
	$fatalError .= '<br /><br /><a href="logupload.htm">Try again</a>';
    exit($fatalError);
}





$templateVersion='';
$templateError='';
$arrayAcceptableTemplates='';
$arrayOldTemplates='';
$arrayNewTemplates='';
$displayBlock = "";
$submitCall='';
$submitName='';
$submitStreet='';
$submitCity='';
$submitStProv='';
$submitCountry='';
$submitPostal='';
$submitZone='';
$submitEmail='';
$submitClub='';
$submitClass='';
$cellValue='';
$errorMessages='';
$qsoErrorMessages='';
$classUnlimited='';
$classLimited='';
$classFormula100='';
$classFormula5='';
$arrayClasses='';
$antennaDesc='';
$swlEntry='';
$currentList='';


//add the CQ header to the display block, so it's there no matter what we return to the user.
$displayBlock .= '<div align="center"><img border="0" src="images/CQ%20logo.jpg" width="276" height="64"></div>';
$displayBlock .= "<h1>CQ Magazine's Annual DX Marathon</h1>";

//get the log file and work with it.
//move_uploaded_file($_FILES['fileLog']["tmp_name"],"uploads/".$logFile);
$logFile = basename($_FILES["fileLog"]["name"]);
$targetPath = "/homepages/41/d154346020/htdocs/dxmarathon/uploads/";
move_uploaded_file($_FILES['fileLog']["tmp_name"],$targetPath.$logFile);
$fullFileName = $targetPath.$logFile;

 // process the file uploaded
 $objReader=PHPExcel_IOFactory::createReaderForFile($fullFileName);
 $objReader->setReadDataOnly(true);
 $objXLS=$objReader->load($fullFileName);	

 //Validate that a proper template was used
 $arrayAcceptableTemplates = array(  //list of acceptable templates
  "2022.1"
  );
 $arrayOldTemplates = array(  //list of old templates (no longer acceptable)
  "2017.1",
  "2018.1",
  "2019.1",
  "2019.2",
  "2020.1",
  "2021.1"
  );
 $arrayNewTemplates = array(  //list of new templates (not yet acceptable)
  "2022.1"
  );

 $templateVersion = $objXLS->getSheet(0)->getCell('B1')->getValue();  //Template version_compare
 if (!in_array($templateVersion,$arrayAcceptableTemplates)) {
 	$objXLS->disconnectWorksheets();
	unset($objXLS);
	unlink($fullFileName);

	 // provide specific reasoning if using an old or new template
	 if (in_array($templateVersion, $arrayOldTemplates)) {
	 	$templateError = "Old template is no longer acceptable";
	 }
	 if (in_array($templateVersion, $arrayNewTemplates)) {
	 	$templateError = "New template is not yet acceptable";
	 }

	$fatalError = '';
	$fatalError .= 'Unable to process this template. ('.$templateError.')<br /><br />';
	$fatalError .= 'Please download the appropriate submission form at <a href="submission.htm#Submission%20Forms">http://www.dxmarathon.com/submission.htm</a';
  exit($fatalError); //we want to stop here and abort if template is not valid.
 }
 
 //Collect and validate the submitter's information
 $submitCall = trim(strtoupper($objXLS->getSheet(0)->getCell('B5')->getValue()), " \t\n\r");  //Call
	if ($submitCall == "") {
		$errorMessages .= "Your submission does not include your callsign. <br />";
	}
	//check for any special characters that have no business being in a callsign
	$pattern = "/[!@#$%^&*()+=_\-\+={}\[\]:;'<>,.?~` \\\\]/";
	if(preg_match($pattern, $submitCall)) {
		$errorMessages .= "Your callsign contains invalid characters. <br />";
	}
 $submitName = trim($objXLS->getSheet(0)->getCell('C5')->getValue(), " \t\n\r");  //Name
	if ($submitName == "") {
		$errorMessages .= "Your submission does not include your name. <br />";
	}
 $submitStreet = trim($objXLS->getSheet(0)->getCell('D5')->getValue(), " \t\n\r");  //Street
	if ($submitStreet == "") {
		$errorMessages .= "Your submission does not include your street address. <br />";
	}
 $submitCity = trim($objXLS->getSheet(0)->getCell('B7')->getValue(), " \t\n\r");  //City
	if ($submitCity == "") {
		$errorMessages .= "Your submission does not include your city. <br />";
	}
 $submitStProv = trim($objXLS->getSheet(0)->getCell('C7')->getValue(), " \t\n\r");  //State or Province, no additional validation required
 $submitCountry = trim($objXLS->getSheet(0)->getCell('D7')->getValue(), " \t\n\r");  //Country
	if ($submitCountry == "") {
		$errorMessages .= "Your submission does not include your country. <br />";
	}
 $submitPostal = trim($objXLS->getSheet(0)->getCell('G7')->getValue(), " \t\n\r");  //Postal code
 	if ($submitPostal == "") {
 		$errorMessages .= "Your submission does not include your postal code. <br />";
 	}
 $submitZone = trim($objXLS->getSheet(0)->getCell('B9')->getValue(), " \t\n\r");  //CQ Zone
	if ($submitZone == "") {
		$errorMessages .= "Your submission does not include your CQ zone. <br />";
	} else {
	 if ($submitZone < 1 || $submitZone > 40) {
		$errorMessages .= "Your submission contains an invalid CQ zone. <br />";
	 }
	}
 $submitEmail = trim($objXLS->getSheet(0)->getCell('C9')->getValue(), " \t\n\r");  //Email
	if ($submitEmail == "") {
		$errorMessages .= "Your submission does not include your email address. <br />";
	}
 $submitClub = trim($objXLS->getSheet(0)->getCell('D9')->getValue(), " \t\n\r");  //Club
 
 //Determine the class based on the populated fields
 $classUnlimited = trim(strtoupper($objXLS->getSheet(0)->getCell('G10')->getValue()), " \t\n\r");  //Unlimited selected
 $classLimited = trim(strtoupper($objXLS->getSheet(0)->getCell('G11')->getValue()), " \t\n\r");  //Limited selected
 $classFormula100 = trim(strtoupper($objXLS->getSheet(0)->getCell('G12')->getValue()), " \t\n\r");  //Formula 100 selected
 $classFormula5 = trim(strtoupper($objXLS->getSheet(0)->getCell('G13')->getValue()), " \t\n\r");  //Formula 5 selected

 
 //set variable to selected class
 if ($classUnlimited == 'X') {
  $submitClass = 'Unlimited';
 } elseif ($classLimited == 'X') {
  $submitClass = 'Limited';
 } elseif ($classFormula100 =='X') {
  $submitClass = 'Formula100';
 } elseif ($classFormula5 == 'X') {
  $submitClass = 'Formula5';
 } else {
 	$submitClass = '';
 }

 if ($submitClass == "") {
 	$errorMessages .= "Your submission does not include your entry class. (blank) <br />";
 } else {
  // Make sure that only one class is selected
  if (($classUnlimited == 'X') && ($classLimited == 'X' || $classFormula100 == 'X' || $classFormula5 =='X')) {
   $errorMessages .= "You can only enter one (1) class. <br />";
  } elseif (($classLimited == 'X') && ($classFormula100 == 'X' || $classFormula5 =='X')) {
   $errorMessages .= "You can only enter one (1) class. <br />";
  } elseif ($classFormula100 == 'X' && $classFormula5 =='X') {
   $errorMessages .= "You can only enter one (1) class. <br />";
  } else {
   // verify antenna description exists if not Unlimited class
   $arrayClasses = array("Limited","Formula100","Formula5");
   if (in_array($submitClass,$arrayClasses)) {
    //class requires antenna description
    $antennaDesc = $objXLS->getSheet(0)->getCell('B14')->getValue();
    if ($antennaDesc == "") {
 	  $errorMessages .= "Your entry class requires an antenna description. <br />";
    }
   }	
  }
 }
	
 
 //check to make sure there is an actual QSOs listed for country and zone credit ... not blank
 $countryCount = $objXLS->getSheet(0)->getCell('I4')->getOldCalculatedValue();
 $zoneCount = $objXLS->getSheet(0)->getCell('I5')->getOldCalculatedValue();

if (is_numeric($countryCount)) {
 if ($countryCount < 1) {
 	//no QSO listed for country credit
 	$qsoErrorMessages .= 'You do not have a valid QSO for country credit.<br />';
 }
} else {
	$qsoErrorMessages .= 'Country count total is invalid character.<br />';
}
if (is_numeric($zoneCount)) {
 if ($zoneCount < 1) {
 	$qsoErrorMessages .= 'You do not have a valid QSO for zone credit.<br />';
 }
} else {
	$qsoErrorMessages .= 'Zone count total is invalid character.<br />';
}


 
 //Check the QSO data
 $row=17;
 $dayValue='';
 $monValue='';
 $utcValue='';
 $bandValue='';
 $modeValue='';
 $callsign='';
 $callsignNoTrim='';
 $entityName='';
 $entityPrefix='';

 while ($row <= 402) {
 	$dayValue = trim($objXLS->getSheet(0)->getCell('D'.$row)->getValue(), " \t\n\r");
 	$monValue = trim($objXLS->getSheet(0)->getCell('E'.$row)->getValue(), " \t\n\r");
 	$utcValue = trim($objXLS->getSheet(0)->getCell('F'.$row)->getValue(), " \t\n\r");
 	$bandValue = trim($objXLS->getSheet(0)->getCell('G'.$row)->getValue(), " \t\n\r");
 	$modeValue = trim(strtoupper($objXLS->getSheet(0)->getCell('H'.$row)->getValue()), " \t\n\r");
 	$callsign = trim($objXLS->getSheet(0)->getCell('I'.$row)->getValue(), " \t\n\r");
 	$callsignNoTrim = $objXLS->getSheet(0)->getCell('I'.$row)->getValue();
 	$entityName = $objXLS->getSheet(0)->getCell('C'.$row)->getValue();
 	$entityPrefix = $objXLS->getSheet(0)->getCell('B'.$row)->getValue();

 	if ($dayValue != "" || $monValue != "" || $utcValue != "" || $bandValue != "" || $modeValue != "" || $callsign != "") {
 	 if ($dayValue == "" || $monValue == "" || $utcValue == "" || $bandValue == "" || $modeValue == "" || $callsign == "") {
 	  $qsoErrorMessages .= $entityName ."(". $entityPrefix ."): Incomplete QSO information.<br />";
 	 }
 	 if (in_array($monValue, array('2','02')) && $dayValue > 29) {
 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid day for month listed.<br />";
 	 }
 	 if (in_array($monValue, array('4','04','6','06','9','09','11')) && $dayValue > 30) {
 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid day for month listed.<br />";
 	 } 
 	 if ($dayValue != "") {
	 	 if ($dayValue > 31 || $dayValue < 1) {
	 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid day, must be 1-31.<br />";
	 	 }
 	 }
 	 if ($monValue != "") {
	 	 if ($monValue < 1 || $monValue > 12) {
	 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid month, must be 1-12.<br />";
	 	 }
	 }
	 if ($utcValue != "") {
	 	 if ($utcValue > 2359) {
	 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid time, must be 0000-2359 UTC.<br />";
	 	 }
	 	 if (strlen($utcValue) < 4) { //pad the value out to 4 digits
	 	 	switch (strlen($utcValue)) {
	 	 		case '1':
	 	 			$utcValue = '000'.$utcValue;
	 	 			break;
	 	 		case '2':
	 	 			$utcValue = '00'.$utcValue;
	 	 			break;
	 	 		case '3':
	 	 			$utcValue = '0'.$utcValue;
	 	 			break;
	 	 		default:
	 	 			// shouldn't have any exceptions here
	 	 			break;
	 	 	}
	 	 } elseif (substr($utcValue, -2) > 59) {
	 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid time, last 2 digits can't be higher than 59.<br />";
	 	 }
 	 }
 	 if ($bandValue != "") {
	 	 if (!in_array($bandValue, array('02','2','04','4','06','6','10','12','15','17','20','30','40','60','80','160'))) {
	 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid band, must be 2, 4, 6, 10, 12, 15, 17, 20, 30, 40, 60, 80, or 160.<br />";
	 	 }
 	 }
 	 if ($modeValue != "") {
	 	 if (!in_array($modeValue, array('CW','PHONE','DIGITAL','SSB'))) {
	 	 	$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Invalid mode, must be CW, PHONE, SSB, or DIGITAL<br />";
	 	 }
 	 }
 	}

 	//one last check to make sure that if all fields are blank but the callsign has spaces.
 	if ($callsignNoTrim != '') {
 		//this means that there is some type of value, either a call of blank spaces
 		if ($callsign == '') {
 			$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Callsign field contains spaces, please delete any spaces in this field.<br />";
 		}
 		//check for any special characters that have no business being in a callsign
 		$pattern = "/[!@#$%^&*()+=_\-\+={}\[\]:;'<>,.?~` \\\\]/";
 		if(preg_match($pattern, $callsign)) {
 			$qsoErrorMessages .= $entityName ." (". $entityPrefix ."): Callsign field contains illegal characters.<br />";
 		}
 	}

 	$row++;
 }
 
 
 //Display Submitter's information from form, and add any error corrections needed.
 $displayBlock .= "Callsign: <b>". $submitCall . "</b><br />";
 $displayBlock .= "Name: <b>". $submitName . "</b><br />";
 $displayBlock .= "Street: <b>". $submitStreet . "</b><br />";
 $displayBlock .= "City: <b>". $submitCity . "</b><br />";
 $displayBlock .= "State/Province: <b>". $submitStProv . "</b><br />";
 $displayBlock .= "Country: <b>". $submitCountry . "</b><br />";
 $displayBlock .= "Postal Code: <b>". $submitPostal . "</b><br />";
 $displayBlock .= "CQ Zone: <b>". $submitZone . "</b><br />";
 $displayBlock .= "Email: <b>". $submitEmail . "</b> (Confirmation will be sent to this address.)<br />";
 $displayBlock .= "Club: <b>". $submitClub . "</b><br />";
 $displayBlock .= "Class: <b>". $submitClass . "</b><br />";

if ($errorMessages != "") {
	$displayBlock .= "<br />Please correct the following errors in your submission form. <br />";
 	$displayBlock .= $errorMessages;
}
 
//Display any QSO error messages
if ($qsoErrorMessages != "") {
	$displayBlock .= "<hr /><p>Please correct the following problems with your QSO list.</p><br />";
	$displayBlock .= "<p>";
	$displayBlock .= $qsoErrorMessages;
	$displayBlock .= "</p>";
} 

//disconnect from the workbook and kill the object
$objXLS->disconnectWorksheets();
unset($objXLS);


// regular or SWL entry (checked on form = on), used for email subject and submitted logs only.
//$swlEntry = $_POST["chkSWL"];


// If everything is good, we want to submit via email, otherwise let them do it again
if ($errorMessages == "" && $qsoErrorMessages == "") {

	//A quick capture of what the file name will be to display to the user on succesful upload.
	$fileNameforDisplay=str_replace("/", "-", $submitCall).".xls";
	$displayBlock .= "File: <b>". $fileNameforDisplay . "</b><br />";

	//rename the file so that it uses the call sign of the submission as the file name
	//need to swap out a / due to portabel indicator for a hyphen to prevent problems
	$newFileName=str_replace("/", "-", $submitCall).".xls";
	rename($fullFileName, $targetPath.$newFileName);

	//email the file to DXMarathon and copy the user
	$mail = new PHPMailer(true);                              // Passing `true` enables exceptions
	try {
	    //Server settings
	    //$mail->SMTPDebug = 2;                                 // Enable verbose debug output
	    //$mail->isSMTP();                                      // Set mailer to use SMTP
	    $mail->IsSendMail();									//Let's just try and use SendMail instead and see if that works
	    $mail->Host = 'smtp.ionos.com';  // Specify main and backup SMTP servers
	    $mail->SMTPAuth = true;                               // Enable SMTP authentication
	    $mail->Username = 'logrobot@dxmarathon.com';                 // SMTP username
	    $mail->Password = 'ND9G-DXer';                           // SMTP password
	    $mail->SMTPSecure = 'tls';                            // Enable TLS encryption, `ssl` also accepted
	    $mail->Port = 587;                                    // TCP port to connect to

	    //Recipients
	    $mail->setFrom('logrobot@dxmarathon.com', 'DX Marathon Robot');
	    $mail->addAddress('logs@dxmarathon.com');     // Add a recipient
	    //$mail->addAddress('ellen@example.com');               // Name is optional
	    $mail->addReplyTo('k9el@dxmarathon.com', 'DX Marathon Admin');
	    $mail->addCC($submitEmail);
	    //$mail->addBCC('bcc@example.com');

	    //Attachments
	    $mail->addAttachment($targetPath.$newFileName);         // Add attachments

	    //Content
	    $mail->isHTML(true);                                  // Set email format to HTML
	    $mail->Subject = $submitCall .' DX Marathon submission';
	    if ($swlEntry = "on") {
	    	$mail->Body    = "Attached is ".$submitName."'s (".$submitCall.") entry for the DX Marathon. (SWL)";
	    } else {
	    	$mail->Body    = "Attached is ".$submitName."'s (".$submitCall.") entry for the DX Marathon.";
	    }
	    $mail->Body    = "Attached is ".$submitName."'s (".$submitCall.") entry for the DX Marathon.";
	    $mail->AltBody = "Attached is ".$submitName."'s (".$submitCall.") entry for the DX Marathon.";

	    $mail->send();
	    //echo 'Message would be sent here, must be a problem with email settings';

	    //the mail has been sent, add the call to the logs received list
	    /*if ($swlEntry == "on") {  // this is an SWL entry
	    	//let's check to make sure it isn't already in the list before adding it again.
	    	$currentList = file('logsSWL.txt', FILE_IGNORE_NEW_LINES); //put existing list in array
	    	if (!in_array($submitCall, $currentList)) { //the submitter's call isn't present
		    	file_put_contents('logsSWL.txt', $submitCall.PHP_EOL , FILE_APPEND | LOCK_EX); //add it to the file
		    }
	    } else { //this is a regular entry*/
	    	//let's check to make sure it isn't already in the list before adding it again.
	    	$currentList = file('logsRcvd.txt', FILE_IGNORE_NEW_LINES); //put existing list in array
	    	if (!in_array($submitCall, $currentList)) { //the submitter's call isn't present
	    		file_put_contents('logsRcvd.txt', $submitCall.PHP_EOL , FILE_APPEND | LOCK_EX);
	    	}
	    // }

	} catch (Exception $e) { 
	    echo 'Message could not be sent. Mailer Error: ', $mail->ErrorInfo;
	    echo '<br />If this problem persists, please contact the DX Marathon administrator at <i><font color="blue">k9el  at dxmarathon.com</font></i>';
	}

	$displayBlock .= '<p>&nbsp;</p><p>&nbsp;</p><p><b>Congratulations</b>, your log has been submitted.</p>';
	$displayBlock .= '<p>A copy has been sent to the email address listed on your application. ';
	$displayBlock .= 'You may check the <a href="submissionlist.php">list of received logs</a> for your call. If you do not see your call listed, ';
	$displayBlock .= 'please email <i><font color="blue">k9el at dxmarathon.com</font></i> to confirm your log was received before uploading again.</p>';

	//delete the file from the server, disabled 08-Nov-2021 to retain files on server.
	//unlink($fullFileName);
} else {
	$displayBlock .= '<br /><p><a href="logupload.htm" class="button-resubmit">Upload Again</a></p>';

	//delete the file from the server
	unlink($fullFileName);
}
 

?>





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
                      "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
	<TITLE>DX Marathon - Log Validation Results</TITLE>
	<META http-equiv="Content-Type" content="text/html; charset=utf-8">
	<STYLE TYPE="text/css">
	 p {
	 	font: 14px ariel, sans-serif;
	 }
	 .button-resubmit {
      background-color: #f44336;; /* Red */
      border: none;
      color: white;
      padding: 15px 32px;
      border-radius: 12px;
      text-align: center;
      text-decoration: none;
      display: inline-block;
      font-size: 16px;
      margin: 4px 30px;
	 }
	 h1 {
		font: 32px ariel, sans-serif;
		text-align: center;
	 }
	</STYLE>
</HEAD>
<BODY>
 <p>&nbsp;</p>
 <p><?php echo $displayBlock ?></p>
</BODY>
</HTML>