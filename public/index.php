<?php
require_once '../vendor/autoload.php';

include '../../phpword/samples/Sample_Header.php';


use PhpOffice\PhpWord\Shared\Converter;

// New Word document
echo date('H:i:s'), ' Create new PhpWord object';

$phpWord = new \PhpOffice\PhpWord\PhpWord();

// 2D charts
$section = $phpWord->addSection();

$section->addTitle(htmlspecialchars('Chart: Basic 2D'), 1);
$section = $phpWord->addSection(array('colsNum' => 2, 'breakType' => 'continuous'));

$chartTypes  = array('pie', 'doughnut', 'bar', 'column', 'line', 'area', 'scatter', 'radar');
$twoSeries   = array('bar', 'column', 'line', 'area', 'scatter', 'radar');
$threeSeries = array('bar', 'line');
$categories  = array('A', 'B', 'C', 'D', 'E');
$series1     = array(1, 3, 2, 5, 4);
$series2     = array(3, 1, 7, 2, 6);
$series3     = array(8, 3, 2, 5, 4);

// Testing Single Chart

$categories = array('F', 'G', 'H', 'I');
$values1    = array(8.2, 3.2, 1.4, 1.2);
$layout     = new \PhpOffice\PhpWord\Writer\Word2007\Part\ChartLayout();
// Testing added chartLayout
$layout->setShowPercent(true);
$layout->setShowSerName(true);

$chart = $section->addChart('pie', $categories, $values1);
$chart->getStyle()->setWidth(Converter::inchToEmu(2.5))->setHeight(Converter::inchToEmu(2));
//$chart->addLegend('l');
$section->addTextBreak();


// Testing Single Chart - 3D

$categories = array('F', 'G', 'H', 'I');
$values1    = array(0.005, 0.003, 0.002, 0.001, 0.008);
$values2    = array(0.005, 0.003, 0.002, 0.001, 0.008);
$chart      = $section->addChart('pie', $categories, $values1);
$chart->addSeries($categories, $values2);
$chart->getStyle()->setWidth(Converter::inchToEmu(2.5))->setHeight(Converter::inchToEmu(5))->set3d(true);
$section->addTextBreak();

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('helloWorld.docx');
