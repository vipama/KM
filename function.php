<?php 
function connect_db(){

	$linkdb = mysqli_connect("localhost","root","j1o2k3e4","mydatabase");
	if ($linkdb == false) {
		# code...
		die("Can not connect DB!");
	} else {
		# code...
		mysqli_query($linkdb,"set names utf8");
		return $linkdb;
	}
	
}

function getAgency_Name($Id)
{
	$link = connect_db();
	$sql = "select ag_ID,ag_Name,ag_Rate,ag_Image from Tb_Agency where ag_ID=".$Id;
	$result = mysqli_query($link,$sql);
	$rows = mysqli_fetch_array($result);
	$heart = "";
	for ($i=0; $i < $rows["ag_Rate"] ; $i++) { 
		if ($rows["ag_Rate"] >= 5) {
			$heart = $heart."<i class='fa fa-heart red'></i>";
		} else {
			$heart = $heart."<i class='fa fa-heart green'></i>";
		}
	}
	//return "&nbsp;".$heart."&nbsp;".$rows["ag_Name"];
	return "&nbsp;".getStar($Id)."&nbsp;".$rows["ag_Name"]."&nbsp;"."<img src='images/".$rows["ag_Image"]."' height='35px' ></img>";
}
function getStar($Id){
	$link = connect_db();
	$sql = "select  ag_Rate from tb_Agency where ag_ID=".$Id;
	$result = mysqli_query($link,$sql);
	$rows = mysqli_fetch_array($result);
	$heart = $rows["ag_Rate"];
	$showheart="";
	for ($i=0; $i < $heart ; $i++) { 
		if ($heart >= 5) {
			$showheart = $showheart."<i class='fa fa-heart red'></i>";
		} else {
			$showheart = $showheart."<i class='fa fa-heart green'></i>";
		}	
	}
	return $showheart;
}
 ?>
