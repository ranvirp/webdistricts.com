<?php

function excel_import_multiple_sheets_fast(&$form_state,$nid) {

if (!user_access('manage excel imports')) {
 return 'You cannot import from excel files';
}
if ($nid){
$excel_import_task=node_load($nid);}
else {
$excel_import_task=$form_state['storage']['excel_import_task'];
}

//$outstr = $t->convert($outstr, $fontface);
 

$flag=0;
$nodecount=0;

$allrows=$excel_import_task->field_excelimportrows[0]['value'];

//dpm($excelformat);
#$sheet_no=$excel_import_task->field_excelimportsheetno[0]['value'];
$status=$excel_import_task->excelimportstatus[0]['value'];
if ($status) { return "Import is already complete";};
$types=$excelformat['type'];
unset($excelformat['type']);
$header[]='Sheet No';
$header[]='Row No';

if (!isset($form_state['storage']['rownos'])){
//new version 
$flag=1;
$matches=array();
//////dsm($allrows);
preg_match_all("/(\d+):((\d+(\-(\d+))?)\,?)+;/",$allrows,$matches);
//////dsm($matches);
foreach ($matches[0] as $key=>$value) {
 preg_match_all("/\,?(\d+)(\-(\d+))?[^:]/",$value,$matches1,PREG_SET_ORDER);
 $sheet_no=$matches[1][$key];
 //////dsm($matches1);

$rownos=array();
foreach ($matches1 as $match) {
if(count($match) ==4) {//row1-row2
  $row1= $match[1];
  $row2=$match[3];
  if ($row1>$row2) { $temp=$row2;$row2=$row1;$row1=$temp;} //swap
  
  $rownos=array('rowbegin'=>$row1,'rowend'=>$row2);
  ////dsm($rownos);
  /*
   for($i=$row1;$i<=$row2;$i++) {
    //check if $i is already in array 
	if (! in_array($i,$rownos) ) {
	  $rownos[]=$i;
	 //$rownos[]=array('sheet'=>$sheet_no,'rowno'=>$i);
	}
   }
   */
   
}
if (count($match)==2) {
 //if (!in_array($match[1],$rownos)) {
   //$rownos[]=$match[1];
   $rownos[]=array('rowbegin'=>$match[1],'rowend'=>$match[1]);
   //$rownos[]=array('sheet'=>$sheet_no,'rowno'=>$match[1]);
 //}
}
//////dsm($rownos);
$form_state['storage']['rownos'][$sheet_no][]=$rownos;
$rownos=array();
}

}


}

$no_of_sheets=count($form_state['storage']['rownos']);
$sheets=array_keys($form_state['storage']['rownos']);
$index1=1;
$sheetindex=0;
$sheetentryindex=0;
$flag=0;
$sheet_no=0;
while (($index1<=50) && ($flag==0)) {
$sheet_no=$sheets[$sheetindex];
 $row_begin=$excel_import_task->field_excelimportcurrentrow[$sheet_no]['value']+1;

 $rowbegin=$form_state['storage']['rownos'][$sheet_no][$sheetentryindex]['rowbegin'];
 if ($row_begin > $rowbegin) {
 drupal_set_message("initial row changed from ".$rowbegin."  to ".$row_begin);
 $rowbegin=$form_state['storage']['rownos'][$sheet_no][$sheetentryindex]['rowbegin']=$row_begin;
 
 }
 //////dsm($rowbegin);
 $rowend=$form_state['storage']['rownos'][$sheet_no][$sheetentryindex]['rowend'];
 ////dsm($rowbegin.".".$rowend);
 if ($rowbegin>$rowend){
 
 if (isset($form_state['storage']['rownos'][$sheet_no][$sheetentryindex+1])){
	$sheetentryindex++;
   } elseif ($sheetindex!=$no_of_sheets-1){
     $sheetindex++;
   } else {
     $flag=1;//we are done
   }
 } else{
  
$index1++;

 
 
 $i=$form_state['storage']['rownos'][$sheet_no][$sheetentryindex]['rowbegin'];
 $form_state['storage']['rownos'][$sheet_no][$sheetentryindex]['rowbegin'] +=1;
 //$i=array_shift($form_state['storage']['rownos'][$sheet_no]);
 $rows[]=$sheet_no;
//return $data->dump();
$rows[]=$i;
}
$rowss[]=$rows;
$rows=array();
}


$form_state['storage'] = isset($form_state['storage']) ? $form_state['storage'] : array();
$form_state['values'] = isset($form_state['values']) ? $form_state['values'] : array();
$values = array_merge($form_state['storage'], $form_state['values']);
$form_state['storage']['preview']=$rowss;

$form_state['storage']['excel_import_task']=$excel_import_task;

$form=array();
#$form['nid']['#value']=$nid;

$form['submit']=array(
'#type'=>'submit',
'#value'=>'submit',
'#submit'=>array('excel_import_multiple_sheets_fast_submit'),
);
$form['#theme']='excel_import_preview';
$form['#preview']=$rowss;
$form['#headers']=$header;
return $form;
}

function excel_import_multiple_sheets_fast_submit($form,&$form_state){


$excel_import_task=$form_state['storage']['excel_import_task'];

$excelfiles=$excel_import_task->field_excelimportfile[0]['filepath'];
$excelfiles=realpath(".")."/".$excelfiles;
$excelformat=unserialize($excel_import_task->field_importmapping[0]['value']);
require_once 'excel_reader2.php';
$data = new Spreadsheet_Excel_Reader($excelfiles);

unset($excelformat['type']);


include_once (dirname(__FILE__) . "/unigateway/Encoder/transformer.php5");
$t = new Transformer();
$outstr="ftyk ;kstuk";
$fontface="kruti dev 010";

foreach ($form_state['storage']['preview'] as $arrayindex=>$row_sheetarray){
$sheet_no=$row_sheetarray[0];
$i=$row_sheetarray[1];
$worksnode=new stdClass();
$worksnode->type=$excel_import_task->field_type[0]['value'];
global $user;
$worksnode->uid=$user->uid;
foreach ($excelformat as $key=>$value) {
	if (is_numeric($value)) {
		$k=$data->val($i,$value,$sheet_no);
		if (($k=="'")||($k=="")) { }
					else { 
						
						$k = htmlentities($k,ENT_NOQUOTES);
						//print "<b>".$k."</b><br>";
						$k=_excel_import_convertLatin1ToHtml($k);
						//print "<b>".$k."</b><br>";
						//$k=encode($k,"UTF-8");
						//mb_convert_encoding($k,"ISO-8859-15");
						//$k=bin2hex($k);
						/*
						$link = $data->hyperlink($i,$value,$sheet_no);
						if ($link!='') {
							$k = "<a href=\"$link\">$k</a>";
						}
						*/
						//$k=convert_uudecode($k);
					//////dsm($k);
			//$k=nl2br($k);

						}
				//////dsm($data->val($i,$value,$sheet_no));
		} 
	else {
		$k=$value;
	}
	$font=$data->font($i,$value,$sheet_no);//debug
	//$font="kruti dev 011";
	if (strtolower($font)==$fontface) {
		$k=$t->convert($k,$font);
	}
	//$rows[]=$k;
	$worksnode->$key=array(array('value'=>$k));
	
}
$panchayat=$worksnode->field_workspanchayat[0]['value'];
$block=$worksnode->field_worksblock[0]['value'];
$revvillage=$worksnode->field_worksrevvill[0]['value'];
$tehsil=$worksnode->field_workstehsil[0]['value'];
$dept=$worksnode->field_worksdept[0]['value'];
$scheme=$worksnode->field_worksscheme[0]['value'];
//check if panchayat/ block exists
//////dsm("Trying to match block".$block." Panchayat:".$panchayat);
		$results=db_query("SELECT node.nid AS nid,
   node_data_field_block.field_block_value AS node_data_field_block_field_block_value,
   node.type AS node_type,
   node.vid AS node_vid,
   node_data_field_block.field_panchayat_value AS node_data_field_block_field_panchayat_value,
   node.uid AS node_uid,
   node_revisions.format AS node_revisions_format
 FROM node node 
 LEFT JOIN content_type_panchayats node_data_field_block ON node.vid = node_data_field_block.vid
 LEFT JOIN node_revisions node_revisions ON node.vid = node_revisions.vid
 WHERE (node.type in ('panchayats')) AND (node_data_field_block.field_panchayat_value = '%s') 
 AND (node_data_field_block.field_block_value = '%s')",$panchayat,$block);
 $count=0;$nidtemp=$nidtemp1=NULL;
 while ($result=db_fetch_object($results)) {
	$count+=1;
	//dd($block."-".$panchayat." vs ".$result->node_data_field_block_field_block_value."-".$result->node_data_field_block_field_panchayat_value);
	$nidtemp=$result->nid;
 }
 if ($count!=1) {
   //dd(" Error ...".$count);
   $panchayatpart=mb_substr($panchayat,0,2);
   //dd($panchayatpart);
   $panchayatlast=mb_substr($panchayat,mb_strlen($panchayat)-2,2);
   //dd("last part=".$panchayatlast);
   $results1=db_query("SELECT node.nid AS nid,
   node_data_field_block.field_block_value AS node_data_field_block_field_block_value,
   node.type AS node_type,
   node.vid AS node_vid,
   node_data_field_block.field_panchayat_value AS node_data_field_block_field_panchayat_value,
   node.uid AS node_uid,
   node_revisions.format AS node_revisions_format
 FROM node node 
 LEFT JOIN content_type_panchayats node_data_field_block ON node.vid = node_data_field_block.vid
 LEFT JOIN node_revisions node_revisions ON node.vid = node_revisions.vid
 WHERE (node.type in ('panchayats')) AND (node_data_field_block.field_panchayat_value like '%s%') AND (node_data_field_block.field_panchayat_value like '%%%s')",$panchayatpart,$panchayatlast);
 //////dsm("likely matches are :");
 //dd("likely matches are :");
 $count1=0;
 while ($result1=db_fetch_object($results1)) {
	$count1+=1;
	//////dsm($block."-".$panchayat." vs ".$result1->node_data_field_block_field_panchayat_value." ".$result1->node_data_field_block_field_block_value);
	//dd($block."-".$panchayat." vs ".$result1->node_data_field_block_field_panchayat_value." ".$result1->node_data_field_block_field_block_value);
	$nidtemp1=$result1->nid;
 }
 if ($count1==1) {
  //dd('hit nid='.$nidtemp1);
  $worksnode->field_workspanchayatref[0]['nid']=$nidtemp1;
 } else {
  //dd($sheetno.":".$rowno."-"."(Panchayat:".$block."--".$panchayat." )");
 }
   
 } else {
 $worksnode->field_workspanchayatref[0]['nid']=$nidtemp;
   //dd("hit nid=".$nidtemp);
 }
 //check for revenue village/tehsil
 
 		$results2=db_query("SELECT node.nid AS nid, node_data_field_revvillage.field_tehsil_value AS 
		node_data_field_revvillage_field_tehsil_value, node.type AS node_type, node.vid AS node_vid, 
		node_data_field_revvillage.field_revvillage_value AS node_data_field_revvillage_field_revvillage_value,
 node.uid AS node_uid, node_revisions.format AS node_revisions_format, node.title AS node_title FROM node node 
 LEFT JOIN content_type_revvillage node_data_field_revvillage ON node.vid = node_data_field_revvillage.vid 
 LEFT JOIN node_revisions node_revisions ON node.vid = node_revisions.vid 
		WHERE (node.type in ('revvillage')) AND (node_data_field_revvillage.field_revvillage_value = '%s') AND 
		(node_data_field_revvillage.field_tehsil_value = '%s')",$revvillage,$tehsil);
 $count=0;$nidtemp=$nidtemp1=NULL;
 while ($result2=db_fetch_object($results2)) {
	$count+=1;
	//dd($block."-".$panchayat." vs ".$result->node_data_field_block_field_block_value."-".$result->node_data_field_block_field_panchayat_value);
	$nidtemp=$result2->nid;
 }
 
 if ($count!=1) {
   //dd(" Error ...".$count);
   $revvillagepart=mb_substr($revvillage,0,2);
   //dd($panchayatpart);
   $revvillagelast=mb_substr($revvillage,mb_strlen($revvillage)-1,1);
   //dd("last part=".$panchayatlast);
   $results3=db_query("SELECT node.nid AS nid, node_data_field_revvillage.field_tehsil_value AS 
		node_data_field_revvillage_field_tehsil_value, node.type AS node_type, node.vid AS node_vid, 
		node_data_field_revvillage.field_revvillage_value AS node_data_field_revvillage_field_revvillage_value,
 node.uid AS node_uid, node_revisions.format AS node_revisions_format, node.title AS node_title FROM node node 
 LEFT JOIN content_type_revvillage node_data_field_revvillage ON node.vid = node_data_field_revvillage.vid 
 LEFT JOIN node_revisions node_revisions ON node.vid = node_revisions.vid 
		WHERE (node.type in ('revvillage')) AND (node_data_field_revvillage.field_revvillage_value like '%s%') AND 
		(node_data_field_revvillage.field_revvillage_value = '%%%s')",$revvillagepart,$revvillagelast);
 //////dsm("likely matches are :");
 //dd("likely matches are :");
 $count1=0;
 while ($result3=db_fetch_object($results3)) {
	$count1+=1;
	//////dsm($block."-".$panchayat." vs ".$result1->node_data_field_block_field_panchayat_value." ".$result1->node_data_field_block_field_block_value);
	//dd($block."-".$panchayat." vs ".$result1->node_data_field_block_field_panchayat_value." ".$result1->node_data_field_block_field_block_value);
	$nidtemp1=$result3->nid;
 }
 if ($count1==1) {
  //dd('hit nid='.$nidtemp1);
  $worksnode->field_worksrevvillref[0]['nid']=$nidtemp1;
 } else {
  //dd($sheetno.":".$rowno."-"."(Rev.village:".$tehsil."--".$revvillage.")..");
 }
   
 } else {
  $worksnode->field_worksrevvillref[0]['nid']=$nidtemp;
   //dd("hit nid=".$nidtemp);
 }
 
 
 //for scheme 
 		$results2=db_query("SELECT node.nid AS nid,
   node.title AS node_title,
   
   node_data_field_scheme.field_scheme_value AS node_data_field_scheme_field_scheme_value,
   node.uid AS node_uid,
   node_revisions.format AS node_revisions_format
 FROM node node 
 
 LEFT JOIN content_type_schemes node_data_field_scheme ON node.vid = node_data_field_scheme.vid
 LEFT JOIN node_revisions node_revisions ON node.vid = node_revisions.vid
 WHERE (node.type in ('schemes'))  AND (node_data_field_scheme.field_scheme_value = '%s')",$scheme);
 $count=0;$nidtemp=$nidtemp1=NULL;
 while ($result2=db_fetch_object($results2)) {
	$count+=1;
	//dd($block."-".$panchayat." vs ".$result->node_data_field_block_field_block_value."-".$result->node_data_field_block_field_panchayat_value);
	$nidtemp=$result2->nid;
 }
 if ($count!=1) {
   //dd(" Error ...".$count);
   
  //dd($sheetno.":".$rowno."-"."(Scheme:".$scheme.")");
 
   
 } else {
  $worksnode->field_worksschemeref[0]['nid']=$nidtemp;
   //dd("hit nid=".$nidtemp);
 }
 
 
 //for dept
 
 		$results2=db_query("SELECT node.nid AS nid,
   node_data_field_deptname.field_deptname_value AS node_data_field_deptname_field_deptname_value,
   node.type AS node_type,
   node.vid AS node_vid
 FROM node node 
 LEFT JOIN content_type_department node_data_field_deptname ON node.vid = node_data_field_deptname.vid
 WHERE (node.type in ('department')) AND (node_data_field_deptname.field_deptname_value = '%s')",$dept);
 $count=0;$nidtemp=$nidtemp1=NULL;
 while ($result2=db_fetch_object($results2)) {
	$count+=1;
	//dd($block."-".$panchayat." vs ".$result->node_data_field_block_field_block_value."-".$result->node_data_field_block_field_panchayat_value);
	$nidtemp=$result2->nid;
 }
 if ($count!=1) {
   //dd(" Error ...".$count);
   
 // dd($sheetno.":".$rowno."-"."(Dept:".$dept.")");
 }
   
  else {
  $worksnode->field_worksagency[0]['nid']=$nidtemp;
   //////dsm("hit nid=".$nidtemp);
 }
 
 
  node_validate($worksnode);
  if (!($errors=form_get_errors())) {
	//$importednode=node_submit($worksnode);
	node_save($worksnode);
	//dsm($worksnode);
	$nodecount++;
	//////dsm($worksnode);
  }
   if ($worksnode->nid){
	
	$excel_import_task->field_excelimportcurrentrow[$sheet_no]['value']=$i;
	dsm($sheet_no."-".$i);
	
  }
	
  
  }


 
 node_save($excel_import_task);
 $form_state['storage']['excel_import_task']=$excel_import_task;
//dsm($excel_import_task);

drupal_set_message($nodecount." no of nodes of type ".$worksnode->type." created. last nid is ".l($worksnode->nid,'node/'.$worksnode->nid));

#$form_state['rebuild']=TRUE;

}