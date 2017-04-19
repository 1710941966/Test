<?php
header("Content-type: text/html; charset=utf-8");
date_default_timezone_set('Asia/Shanghai');
include_once 'Classes/PHPExcel.php';

// $path="C:\Users\martincui\Desktop\Malaysia Postal Code (duplicates removed).xlsx";

// $fileinfo=pathinfo($path);

// $extension=$fileinfo['extension'];

// if( $extension =='xlsx' ){
// $PHPReader = new PHPExcel_Reader_Excel2007();
// }else{
// $PHPReader = new PHPExcel_Reader_Excel5();
// }

// $PHPExcel=$PHPReader->load($path);

// $currentSheet = $PHPExcel->getSheet(0);

// $allRow = $currentSheet->getHighestRow();

// $data='';

// for ($i=2;$i<=$allRow;$i++){
// $data.="INSERT INTO postalcode (country,countrycode,postalcode,Status,adddate,lastmodifieddate) VALUES('Malaysia','MYS','".$currentSheet->getCell('A'.$i)->getValue()."',1,".time().",".time().");<br/>";
// }

// print_r($data);

include_once '../fwadmin_lifung_splitorder/lib/fw_db.php';
$dbhost = '127.0.0.1';
$dbuser = 'root';
$dbpw = 'root';
$dbname = 'fireswirlstore_jockey';


// $dbhost = '172.26.26.2';
// $dbuser = 'jockey';
// $dbpw = 'FSWsz1403';
// $dbname = 'fireswirlstore_jockey';


$db = new fw_db($dbhost, $dbuser, $dbpw, $dbname);


$sku='JJ13642115150M';
$InventorySql = "select `ProductId`,`OptionToValueList` from `producttodetail` where `PoductInventoryCoding`='" . $sku . "' LIMIT 0,1";
$subproArr = $db->row_query($InventorySql);

print_r($subproArr);exit();




$sql = "select p.ProductId,p.Sku as PLU, pdvl.ProductDetailOptionValueId, pdv.ProductDetailOptionValueCode, pdvl.ProductDetailOptionValue, pimg.ProductImage,
pde.ProductName as ProductEngName, pdsc.ProductName as ProductSCName, pdtc.ProductName as ProductTCName 
from products p left join (select * from productsimages where ProductImageOrder=1) pimg on p.productid = pimg.productid
left join (select * from producttodetailsearch where ProductOptionDetailId=12) pds on p.productid = pds.productid
left join (select * from productdetailoptionvaluelanguage where languageId = 1) pdvl on pds.ProductDetailValueId = pdvl.ProductDetailOptionValueId
left join productdetailoptionsvalue pdv on pdvl.ProductDetailOptionValueId = pdv.ProductDetailOptionValueId
left join (select * from productsdescription where LanguageId = 1) pde on p.ProductId = pde.ProductId
left join (select * from productsdescription where LanguageId = 2) pdsc on p.ProductId = pdsc.ProductId
left join (select * from productsdescription where LanguageId = 3) pdtc on p.ProductId = pdtc.ProductId WHERE p.ProductTypeId in(0,1);";

$data = $db->row_query($sql);

// 创建excel表生成类
$objPHPExcel = new PHPExcel();
$reportSnArr = array(
    'A' => 'Seq.',
    'B' => 'Product ID',
    'C' => 'PLU',
    'D' => 'Id',
    'E' => '属性值code',
    'F' => 'Color',
    'G' => 'ProductImage',
    'H' => '产品名称（sc）',
);
$objPHPExcel->setActiveSheetIndex(0);
$objActSheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->getActiveSheet()->setTitle('PLP');

$objPHPExcel->setActiveSheetIndex(0)->getStyle('A1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('B1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('C1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('D1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('E1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('F1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('G1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(0)->getStyle('H1')->getFont()->setBold(true);

foreach ($reportSnArr as $key => $val) {
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($key . '1', $val);
    $objActSheet->getColumnDimension($key)->setAutoSize(true); // 自动设置单元格宽度
}

for ($i = 0; $i < count($data); $i ++) {
    $n = $i + 2;
    $productId = $data[$i]['ProductId'];
    $plu = $data[$i]['PLU'];
    $colorId = $data[$i]['ProductDetailOptionValueId'];
    $colorCode = htmlspecialchars_decode($data[$i]['ProductDetailOptionValueCode'],ENT_QUOTES);
    $colorName = htmlspecialchars_decode($data[$i]['ProductDetailOptionValue'],ENT_QUOTES);
    $productImg = $data[$i]['ProductImage'];
    $productName = $data[$i]['ProductSCName'];
    
    $objActSheet->setCellValue('A' . $n, ' ' . $i + 1);
    $objActSheet->setCellValue('B' . $n, $productId);
    $objActSheet->setCellValue('C' . $n, $plu);
    $objActSheet->setCellValue('D' . $n, $colorId);
    $objActSheet->setCellValue('E' . $n, $colorCode);
    $objActSheet->setCellValue('F' . $n, $colorName);
    $objActSheet->setCellValue('G' . $n, $productImg);
    $objActSheet->setCellValue('H' . $n, $productName);
}

// 创建第二个工作表

$reportSnArr = array(
    'A' => 'Seq.',
    'B' => 'Product ID',
    'C' => 'PLU',
    'D' => 'Id',
    'E' => '属性值code',
    'F' => 'Color',
    'G' => 'ProductImage',
    'H' => '产品名称（sc）',
    'I' => '添加时间'
);

$msgWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'ShopCart'); // 创建一个工作表
$objPHPExcel->addSheet($msgWorkSheet); // 插入工作表
$objPHPExcel->setActiveSheetIndex(1); // 切换到新创建的工作表
$objActSheet2 = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(1)->getStyle('A1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('B1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('C1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('D1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('E1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('F1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('G1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('H1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(1)->getStyle('I1')->getFont()->setBold(true);
foreach ($reportSnArr as $key => $val) {
    $objPHPExcel->setActiveSheetIndex(1)->setCellValue($key . '1', $val);
    $objActSheet2->getColumnDimension($key)->setAutoSize(true); // 自动设置单元格宽度
    $objPHPExcel->getActiveSheet()->getColumnDimension($key)->setAutoSize(true);
}

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);

$sql='SELECT ProductId,AddDate FROM shopcart';

$shopCartTime = $db->row_query($sql);
//print_r($shopCartTime);
if (count($shopCartTime)>0){
    foreach ($shopCartTime as $v){
        $sql1[] = "select p.ProductId,p.Sku as PLU, pdvl.ProductDetailOptionValueId, pdv.ProductDetailOptionValueCode, pdvl.ProductDetailOptionValue, pimg.ProductImage,
pde.ProductName as ProductEngName, pdsc.ProductName as ProductSCName, pdtc.ProductName as ProductTCName
from products p left join (select * from productsimages where ProductImageOrder=1) pimg on p.productid = pimg.productid
left join (select * from producttodetailsearch where ProductOptionDetailId=12) pds on p.productid = pds.productid
left join (select * from productdetailoptionvaluelanguage where languageId = 1) pdvl on pds.ProductDetailValueId = pdvl.ProductDetailOptionValueId
left join productdetailoptionsvalue pdv on pdvl.ProductDetailOptionValueId = pdv.ProductDetailOptionValueId
left join (select * from productsdescription where LanguageId = 1) pde on p.ProductId = pde.ProductId
left join (select * from productsdescription where LanguageId = 2) pdsc on p.ProductId = pdsc.ProductId
left join (select * from productsdescription where LanguageId = 3) pdtc on p.ProductId = pdtc.ProductId WHERE p.ProductTypeId in(0,1) AND p.ProductId in(SELECT ".intval($v['ProductId'])." FROM shopcart)";
    }
}

foreach ($sql1 as $a){
    $shopCart1=$db->row_query($a);
    $shopCart[]=$shopCart1[0];
    
}

for ($z=0;$z<count($shopCart);$z++){
    for ($y=0;$y<count($shopCartTime);$y++){
        if ($z==$y){
            $shopCart[$z]['AddDate']=$shopCartTime[$y]['AddDate'];
        }
    }
}

for ($i = 0; $i < count($shopCart); $i ++) {
    $n = $i + 2;
    $productId = $shopCart[$i]['ProductId'];
    $plu = $shopCart[$i]['PLU'];
    $colorId = $shopCart[$i]['ProductDetailOptionValueId'];
    $colorCode = htmlspecialchars_decode($shopCart[$i]['ProductDetailOptionValueCode'],ENT_QUOTES);
    $colorName = htmlspecialchars_decode($shopCart[$i]['ProductDetailOptionValue'],ENT_QUOTES);
    $productImg = $shopCart[$i]['ProductImage'];
    $productName = $shopCart[$i]['ProductSCName'];
    $addtime= date('Y-m-d H:i:s',$shopCart[$i]['AddDate']);
    
    $objActSheet2->setCellValue('A' . $n, ' ' . $i + 1);
    $objActSheet2->setCellValue('B' . $n, $productId);
    $objActSheet2->setCellValue('C' . $n, $plu);
    $objActSheet2->setCellValue('D' . $n, $colorId);
    $objActSheet2->setCellValue('E' . $n, $colorCode);
    $objActSheet2->setCellValue('F' . $n, $colorName);
    $objActSheet2->setCellValue('G' . $n, $productImg);
    $objActSheet2->setCellValue('H' . $n, $productName);
    $objActSheet2->setCellValue('I' . $n, $addtime);
}

// 创建第三个工作表
$msgWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'Wishlist'); // 创建一个工作表
$objPHPExcel->addSheet($msgWorkSheet); // 插入工作表
$objPHPExcel->setActiveSheetIndex(2); // 切换到新创建的工作表
$objActSheet3 = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(2)->getStyle('A1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('B1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('C1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('D1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('E1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('F1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('G1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('H1')->getFont()->setBold(true);
$objPHPExcel->setActiveSheetIndex(2)->getStyle('I1')->getFont()->setBold(true);
foreach ($reportSnArr as $key => $val) {
    $objPHPExcel->setActiveSheetIndex(2)->setCellValue($key . '1', $val);
    $objActSheet3->getColumnDimension($key)->setAutoSize(true); // 自动设置单元格宽度
    $objPHPExcel->getActiveSheet(2)->getColumnDimension($key)->setAutoSize(true);
}

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);

$sql = "SELECT * FROM customeroptions WHERE CustomerOptionsNme='favorite';";

$favorite = $db->row_query($sql);

foreach ($favorite as $k=>$val) {
    $pro_info = json_decode(base64_decode($val['CusomerOptionsValue']), true);
    $addtime1=$val['AddDate'];
    $productId = $pro_info[count($pro_info) - 1];
    $wishListProduct[$k]['ProductId'] = $productId;
    $wishListProduct[$k]['AddDate'] = $addtime1;
}

if (count($wishListProduct)>0){
    foreach ($wishListProduct as $v){
        $sql2[] = "select p.ProductId,p.Sku as PLU, pdvl.ProductDetailOptionValueId, pdv.ProductDetailOptionValueCode, pdvl.ProductDetailOptionValue, pimg.ProductImage,
pde.ProductName as ProductEngName, pdsc.ProductName as ProductSCName, pdtc.ProductName as ProductTCName
from products p left join (select * from productsimages where ProductImageOrder=1) pimg on p.productid = pimg.productid
left join (select * from producttodetailsearch where ProductOptionDetailId=12) pds on p.productid = pds.productid
left join (select * from productdetailoptionvaluelanguage where languageId = 1) pdvl on pds.ProductDetailValueId = pdvl.ProductDetailOptionValueId
left join productdetailoptionsvalue pdv on pdvl.ProductDetailOptionValueId = pdv.ProductDetailOptionValueId
left join (select * from productsdescription where LanguageId = 1) pde on p.ProductId = pde.ProductId
left join (select * from productsdescription where LanguageId = 2) pdsc on p.ProductId = pdsc.ProductId
left join (select * from productsdescription where LanguageId = 3) pdtc on p.ProductId = pdtc.ProductId WHERE p.ProductTypeId in(0,1) AND p.ProductId in(SELECT ".intval($v['ProductId'])." FROM shopcart)";
    }
}

foreach ($sql2 as $a){
    $wishList1=$db->row_query($a);
    $wishList[]=isset($wishList1[0])?$wishList1[0]:'';

}

for ($z=0;$z<count($wishList);$z++){
    for ($y=0;$y<count($wishListProduct);$y++){
        if ($z==$y){
            $wishList[$z]['AddDate']=@$wishListProduct[$y]['AddDate'];
        }
    }
}




//$wishList = $db->row_query($sql);

for ($i = 0; $i < count($wishList); $i ++) {
    $n = $i + 2;
    $productId = @$wishList[$i]['ProductId'];
    $plu = @$wishList[$i]['PLU'];
    $colorId = @$wishList[$i]['ProductDetailOptionValueId'];
    $colorCode = htmlspecialchars_decode(@$wishList[$i]['ProductDetailOptionValueCode'],ENT_QUOTES);
    $colorName = htmlspecialchars_decode(@$wishList[$i]['ProductDetailOptionValue'],ENT_QUOTES);
    $productImg = @$wishList[$i]['ProductImage'];
    $productName = @$wishList[$i]['ProductSCName'];
    $addtime= date('Y-m-d H:i:s',$wishList[$i]['AddDate']);
    
    $objActSheet3->setCellValue('A' . $n, ' ' . $i + 1);
    $objActSheet3->setCellValue('B' . $n, $productId);
    $objActSheet3->setCellValue('C' . $n, $plu);
    $objActSheet3->setCellValue('D' . $n, $colorId);
    $objActSheet3->setCellValue('E' . $n, $colorCode);
    $objActSheet3->setCellValue('F' . $n, $colorName);
    $objActSheet3->setCellValue('G' . $n, $productImg);
    $objActSheet3->setCellValue('H' . $n, $productName);
    $objActSheet3->setCellValue('I' . $n, $addtime);
}

$outputFileName = "productsInfo.xlsx";
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save($outputFileName);
chmod($outputFileName, 0777);
header("Content-Disposition: attachment; filename='" . $outputFileName . "'");
header('Content-Type: application/vnd.ms-excel; charset=utf-8');
header('Content-Length: ' . filesize($outputFileName));
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate');
header('Pragma: public');
readfile($outputFileName);
unlink($outputFileName);

