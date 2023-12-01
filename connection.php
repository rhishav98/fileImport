<?php
 $hostname = "localhost";
 $username = "root";
 $password = "";
 $dbname = "fileimport";
 
 $db = new mysqli($hostname,$username,$password,$dbname);

 if($db ->connect_error){
    die("Connection Failes ". $db->connect_error);
 }else{
   
 }
?>