<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

include 'connection.php'; 
require 'vendor/autoload.php';
date_default_timezone_set('Asia/Kolkata');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    if (isset($_POST['submit'])) {
   
        $fileName = $_FILES['import_file']['name'];
      

        $file_ext = pathinfo($fileName, PATHINFO_EXTENSION);
    
        $allowed_ext = ['xls', 'csv', 'xlsx'];



        if (in_array($file_ext, $allowed_ext)) {
            $inputFileNamePath = $_FILES['import_file']['tmp_name'];

            $currDate = date('d-m-y h:i:s');
            $currUser =  get_current_user();
            $userIP = $_SERVER['REMOTE_ADDR'];  
            $impNo = pathinfo($fileName, PATHINFO_FILENAME)."/".$currUser."/".date('d');
           
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileNamePath);
                $worksheet = $spreadsheet->getActiveSheet();
                $data = $worksheet->toArray();

                // Remove header row 
                unset($data[0]);


                foreach ($data as $row) {

                   // Check for null or empty
                            $Project_Creation_Date = isset($row[0]) ? date('ymd', strtotime(mysqli_real_escape_string($db, $row[0]))) : 'NULL';
                            $Target_Date_Of_Completion = isset($row[1]) ? date('ymd', strtotime(mysqli_real_escape_string($db, $row[1]))) : 'NULL';
                                
                            $Project_Name = isset($row[2]) ? mysqli_real_escape_string($db, $row[2]) : 'NULL'; 
                            $Scheme_ID = isset($row[3]) ? mysqli_real_escape_string($db, $row[3]) : 'NULL'; 
                            $Sub_Scheme_ID = isset($row[4]) ? mysqli_real_escape_string($db, $row[4]) : 'NULL'; 
                            $Agency_ID = isset($row[5]) ? mysqli_real_escape_string($db, $row[5]) : 'NULL'; 
                            $Child_Agency_ID = isset($row[6]) ? mysqli_real_escape_string($db, $row[6]) : 'NULL'; 
                            $State_ID = isset($row[7]) ? mysqli_real_escape_string($db, $row[7]) : 'NULL'; 
                            $Sector_ID = isset($row[8]) ? mysqli_real_escape_string($db, $row[8]) : 'NULL'; 

                            $Latitude = isset($row[9]) ? mysqli_real_escape_string($db, $row[9]) : 'NULL'; 
                            $Longitude = isset($row[10]) ? mysqli_real_escape_string($db, $row[10]) : 'NULL'; 
                            $Authorized_Amount = isset($row[11]) ? mysqli_real_escape_string($db, $row[11]) : 'NULL'; 
                            $Drawing_Limit_Amount = isset($row[12]) ? mysqli_real_escape_string($db, $row[12]) : 'NULL'; 
                            $Expenditure_Amount = isset($row[13]) ? mysqli_real_escape_string($db, $row[13]) : 'NULL'; 
                            $Expenditure_As_Per_PFMS = isset($row[14]) ? mysqli_real_escape_string($db, $row[14]) : 'NULL'; 
                            $Balance_Amount = isset($row[15]) ? mysqli_real_escape_string($db, $row[15]) : 'NULL'; 
                            $Total_Project_Cost = isset($row[16]) ? mysqli_real_escape_string($db, $row[16]) : 'NULL'; 
                            $Project_Notes = isset($row[17]) ? mysqli_real_escape_string($db, $row[17]) : 'NULL'; 
                            $Status = isset($row[18]) ? mysqli_real_escape_string($db, $row[18]) : 'NULL'; 
                            $Project_Completion_Date = isset($row[19]) ? mysqli_real_escape_string($db, $row[19]) : 'NULL'; 
                            $Project_Code = isset($row[20]) ? mysqli_real_escape_string($db, $row[20]) : 'NULL'; 
                            $Order_Sanction_No = isset($row[21]) ? mysqli_real_escape_string($db, $row[21]) : 'NULL'; 
                            $Allocation_Type = isset($row[22]) ? mysqli_real_escape_string($db, $row[22]) : 'NULL'; 
                            $Financial_Year = isset($row[23]) ? mysqli_real_escape_string($db, $row[23]) : 'NULL'; 
                            $Agency_Type = isset($row[24]) ? mysqli_real_escape_string($db, $row[24]) : 'NULL'; 
                            $Project_Adrdess = isset($row[25]) ? mysqli_real_escape_string($db, $row[25]) : 'NULL'; 
                            $Unique_ID = isset($row[26]) ? mysqli_real_escape_string($db, $row[26]) : 'NULL'; 


                                // Check  Unique_ID already exists in the database

                            $checkQuery = "SELECT COUNT(*) FROM `projects` WHERE `Unique_ID` = '$Unique_ID'";
                            $result = mysqli_query($db, $checkQuery);
                                // echo "res".$result;
                            $row_count = mysqli_fetch_array($result)[0];
                            

                            if ($row_count == 0) { 
                            $insertQuery = "INSERT INTO `projects` (`Import_No`,
                                `Project_Creation_Date`, `Target_Date_Of_Completion`, `Project_Name`, 
                                `Scheme_ID`, `Sub_Scheme_ID`, `Agency_ID`, `Child_Agency_ID`, `State_ID`,
                                `Sector_ID`, `Latitude`, `Longitude`, `Authorized_Amount`, 
                                `Drawing_Limit_Amount`, `Expenditure_Amount`, `Expenditure_As_Per_PFMS`, 
                                `Balance_Amount`, `Total_Project_Cost`, `Project_Notes`, `Status`, 
                                `Project_Completion_Date`, `Project_Code`, `Order_Sanction_No`, 
                                `Allocation_Type`, `Financial_Year`, `Agency_Type`, `Project_Adrdess`, `Unique_ID`
                                ) VALUES (
                                '$impNo','$Project_Creation_Date', '$Target_Date_Of_Completion', '$Project_Name', 
                                '$Scheme_ID', '$Sub_Scheme_ID', '$Agency_ID', '$Child_Agency_ID', '$State_ID',
                                '$Sector_ID', '$Latitude', '$Longitude', '$Authorized_Amount', 
                                '$Drawing_Limit_Amount', '$Expenditure_Amount', '$Expenditure_As_Per_PFMS', 
                                '$Balance_Amount', '$Total_Project_Cost', '$Project_Notes', '$Status', 
                                '$Project_Completion_Date', '$Project_Code', '$Order_Sanction_No', 
                                '$Allocation_Type', '$Financial_Year', '$Agency_Type', '$Project_Adrdess', '$Unique_ID'
                            )";
                            mysqli_query($db, $insertQuery);      
                        } 
                        else {

                        $updateQuery = "UPDATE `projects` SET
                        `Import_No`='$impNo',
                        `Project_Creation_Date`='$Project_Creation_Date',
                        `Target_Date_Of_Completion`='$Target_Date_Of_Completion',
                        `Project_Name`='$Project_Name',
                        `Scheme_ID`='$Scheme_ID',
                        `Sub_Scheme_ID`='$Sub_Scheme_ID',
                        `Agency_ID`='$Agency_ID',
                        `Child_Agency_ID`='$Child_Agency_ID',
                        `State_ID`='$State_ID',
                        `Sector_ID`='$Sector_ID',
                        `Latitude`='$Latitude',
                        `Longitude`='$Longitude',
                        `Authorized_Amount`='$Authorized_Amount',
                        `Drawing_Limit_Amount`='$Drawing_Limit_Amount',
                        `Expenditure_Amount`='$Expenditure_Amount',
                        `Expenditure_As_Per_PFMS`='$Expenditure_As_Per_PFMS',
                        `Balance_Amount`='$Balance_Amount',
                        `Total_Project_Cost`='$Total_Project_Cost',
                        `Project_Notes`='$Project_Notes',
                        `Status`='$Status',
                        `Project_Completion_Date`='$Project_Completion_Date',
                        `Project_Code`='$Project_Code',
                        `Order_Sanction_No`='$Order_Sanction_No',
                        `Allocation_Type`='$Allocation_Type',
                        `Financial_Year`='$Financial_Year',
                        `Agency_Type`='$Agency_Type',
                        `Project_Adrdess`='$Project_Adrdess'
                        WHERE `Unique_ID`='$Unique_ID'";

                        mysqli_query($db, $updateQuery);
                    }
       }
           
         
           $insertRecord =mysqli_query($db,"INSERT INTO `import_master`( `Import_No`, `File_Name`, `Created_At`,`Created_By`,`User_IP`) 
                                      VALUES ('$impNo','$fileName','$currDate','$currUser','$userIP')");  


       echo '<script> alert("Upload Successful");</script>';
       header('Location: index.php');

   } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo 'Error loading the spreadsheet: ' . $e->getMessage();
}
} else {
    echo '<script> alert("Upload Correct File");</script>';
    header('Location: index.php');
    exit();
}
}
}
