<?php
if(!defined('sugarEntry') || !sugarEntry) die('Not A Valid Entry Point');
/*********************************************************************************
 * SugarCRM Community Edition is a customer relationship management program developed by
 * SugarCRM, Inc. Copyright (C) 2004-2013 SugarCRM Inc.
 * 
 * This program is free software; you can redistribute it and/or modify it under
 * the terms of the GNU Affero General Public License version 3 as published by the
 * Free Software Foundation with the addition of the following permission added
 * to Section 15 as permitted in Section 7(a): FOR ANY PART OF THE COVERED WORK
 * IN WHICH THE COPYRIGHT IS OWNED BY SUGARCRM, SUGARCRM DISCLAIMS THE WARRANTY
 * OF NON INFRINGEMENT OF THIRD PARTY RIGHTS.
 * 
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE.  See the GNU Affero General Public License for more
 * details.
 * 
 * You should have received a copy of the GNU Affero General Public License along with
 * this program; if not, see http://www.gnu.org/licenses or write to the Free
 * Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301 USA.
 * 
 * You can contact SugarCRM, Inc. headquarters at 10050 North Wolfe Road,
 * SW2-130, Cupertino, CA 95014, USA. or at email address contact@sugarcrm.com.
 * 
 * The interactive user interfaces in modified source and object code versions
 * of this program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU Affero General Public License version 3.
 * 
 * In accordance with Section 7(b) of the GNU Affero General Public License version 3,
 * these Appropriate Legal Notices must retain the display of the "Powered by
 * SugarCRM" logo. If the display of the logo is not reasonably feasible for
 * technical reasons, the Appropriate Legal Notices must display the words
 * "Powered by SugarCRM".
 ********************************************************************************/

//Bug 30094, If zlib is enabled, it can break the calls to header() due to output buffering. This will only work php5.2+
ini_set('zlib.output_compression', 'Off');

ob_start();
require_once('export_xls_utils.php');
require_once('custom/include/PHPExcel1.8.0/PHPExcel.php');
global $sugar_config;
global $locale;
global $current_user;
global $app_list_strings;

$the_module = clean_string($_REQUEST['module']);

if($sugar_config['disable_export'] 	|| (!empty($sugar_config['admin_export_only']) && !(is_admin($current_user) || (ACLController::moduleSupportsACL($the_module)  && ACLAction::getUserAccessLevel($current_user->id,$the_module, 'access') == ACL_ALLOW_ENABLED &&
    (ACLAction::getUserAccessLevel($current_user->id, $the_module, 'admin') == ACL_ALLOW_ADMIN ||
     ACLAction::getUserAccessLevel($current_user->id, $the_module, 'admin') == ACL_ALLOW_ADMIN_DEV))))){
	die($GLOBALS['app_strings']['ERR_EXPORT_DISABLED']);
}

//check to see if this is a request for a sample or for a regular export
if(!empty($_REQUEST['sample'])){
    //call special method that will create dummy data for bean as well as insert standard help message.
    $content = exportSample(clean_string($_REQUEST['module']));

}else if(!empty($_REQUEST['uid'])){
	$content = export(clean_string($_REQUEST['module']), $_REQUEST['uid'], isset($_REQUEST['members']) ? $_REQUEST['members'] : false);
}else{
	$content = export(clean_string($_REQUEST['module']));
}
$filename = $_REQUEST['module'];
//use label if one is defined
if(!empty($app_list_strings['moduleList'][$_REQUEST['module']])){
    $filename = $app_list_strings['moduleList'][$_REQUEST['module']];
}

//strip away any blank spaces
$filename = str_replace(' ','',$filename);

//$transContent = $GLOBALS['locale']->translateCharset($content, 'UTF-8', $GLOBALS['locale']->getExportCharset());

if(isset($_REQUEST['members']) && $_REQUEST['members'] == true)
	$filename .= '_'.'members';

$oPhpExcel = new PHPExcel();
$oPhpExcel->getProperties()->setTitle("export")->setDescription("none");
$oPhpExcel->setActiveSheetIndex(0);

// fill color header
$styleArray = array(
    'font'      => array('bold' => true),
    'alignment' => array('horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),
    'borders'   => array('allborders' => array('style' => \PHPExcel_Style_Border::BORDER_THIN)),
    'fill'      => array(
        'type' => \PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
        'rotation'   => 90,
        'startcolor' => array('argb' => 'FFA0A0A0'),
        'endcolor'   => array('argb' => 'FFFFFFFF')
    )
);

$borders = array('borders' => array('allborders' => array('style' => \PHPExcel_Style_Border::BORDER_THIN) ));

$col = 0;

// Set title
//$oPhpExcel->getActiveSheet()->setCellValue('A1', "Đại đẹp trai");
//$oPhpExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//$oPhpExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(16);

//merge cell data
$arrHeader = $content['header'];
//$end = count($arrHeader);
//$end = chr(65 + $end - 1);
//$oPhpExcel->getActiveSheet()->mergeCells('A1:'.$end.'1');
//unset($end);

//set header
foreach ($arrHeader as $field)
{
    $oPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($col, 1, $field);
    $oPhpExcel->getActiveSheet()->getStyleByColumnAndRow($col, 1)->applyFromArray($styleArray);
    $oPhpExcel->getActiveSheet()->getColumnDimensionByColumn($col)->setAutoSize(true);
    $oPhpExcel->getActiveSheet()->getStyleByColumnAndRow($col, 1)->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('999999');
    $col++;
}

$row = 2;
foreach ($content['data'] as $i => $arrData)
{
    //set start column
    $col = 0;
    
    //set cell style
    $oPhpExcel->getActiveSheet()->getStyleByColumnAndRow($col, $row)->applyFromArray($borders);
    $oPhpExcel->getActiveSheet()->getStyleByColumnAndRow($col, $row)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    
    foreach ($arrData as $value)
    {
        $val = html_entity_decode($value, ENT_QUOTES, 'UTF-8');        
         
        $oPhpExcel->getActiveSheet()->getStyleByColumnAndRow($col, $row)->applyFromArray($borders);
        $oPhpExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $val);
    
        $col++;
    }
    $row++;
}

// Download file
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="'.$filename.'_'.date('dMy').'.xls"');
header('Cache-Control: max-age=0');
$oWriter = \PHPExcel_IOFactory::createWriter($oPhpExcel, 'Excel5');
$oWriter->save('php://output');
exit();

///////////////////////////////////////////////////////////////////////////////
////	BUILD THE EXPORT FILE
ob_clean();
header("Pragma: cache");
header("Content-type: application/octet-stream; charset=".$GLOBALS['locale']->getExportCharset());
header("Content-Disposition: attachment; filename={$filename}.csv");
header("Content-transfer-encoding: binary");
header("Expires: Mon, 26 Jul 1997 05:00:00 GMT" );
header("Last-Modified: " . TimeDate::httpTime() );
header("Cache-Control: post-check=0, pre-check=0", false );
header("Content-Length: ".mb_strlen($transContent, '8bit'));
if (!empty($sugar_config['export_excel_compatible'])) {
    $transContent=chr(255) . chr(254) . mb_convert_encoding($transContent, 'UTF-16LE', 'UTF-8');
}
print $transContent;

sugar_cleanup(true);
