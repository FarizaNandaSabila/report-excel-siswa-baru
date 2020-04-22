<?php
include('koneksiform.php');
require ('vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'id');
$sheet->setCellValue('B1', 'Nama Lengkap');
$sheet->setCellValue('C1', 'Jenis Kelamin');
$sheet->setCellValue('D1', 'NISN');
$sheet->setCellValue('E1', 'NIK');
$sheet->setCellValue('F1', 'Tempat Lahir');
$sheet->setCellValue('G1', 'Tanggal Lahir');
$sheet->setCellValue('H1', 'No. Akta');
$sheet->setCellValue('I1', 'Agama');
$sheet->setCellValue('J1', 'Kewarganegaraan');
$sheet->setCellValue('K1', 'Berkebutuhan Khusus');
$sheet->setCellValue('L1', 'Alamat');
$sheet->setCellValue('M1', 'Rt');
$sheet->setCellValue('N1', 'Rw');
$sheet->setCellValue('O1', 'Dusun');
$sheet->setCellValue('P1', 'Desa');
$sheet->setCellValue('Q1', 'Kecamatan');
$sheet->setCellValue('R1', 'Kode pos');
$sheet->setCellValue('S1', 'Lintang');
$sheet->setCellValue('T1', 'Bujur');
$sheet->setCellValue('U1', 'Tempat Tinggal');
$sheet->setCellValue('V1', 'Mode Transport');
$sheet->setCellValue('W1', 'Nomor KKS');
$sheet->setCellValue('X1', 'Anak ke');
$sheet->setCellValue('Y1', 'KPS/KPH');
$sheet->setCellValue('Z1', 'No. KPS/KPH');

$query = mysqli_query($koneksi,"select * from form");
$i = 2;
$no = 1;
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $id);
	$sheet->setCellValue('B'.$i, $row['Nama Lengkap']);
	$sheet->setCellValue('C'.$i, $row['Jenis Kelamin']);
	$sheet->setCellValue('D'.$i, $row['NISN']);	
	$sheet->setCellValue('E'.$i, $row['NIK']);	
	$sheet->setCellValue('F'.$i, $row['Tempat Lahir']);	
	$sheet->setCellValue('G'.$i, $row['Tanggal Lahir']);	
	$sheet->setCellValue('H'.$i, $row['No. Akta']);	
	$sheet->setCellValue('I'.$i, $row['Agama']);
	$sheet->setCellValue('J'.$i, $row['Kewarganegaraan']);	
	$sheet->setCellValue('K'.$i, $row['Berkebutuhan Khusus']);	
	$sheet->setCellValue('L'.$i, $row['Alamat']);	
	$sheet->setCellValue('M'.$i, $row['Rt']);	
	$sheet->setCellValue('N'.$i, $row['Rw']);	
	$sheet->setCellValue('O'.$i, $row['Dusun']);
	$sheet->setCellValue('P'.$i, $row['Desa']);		
	$sheet->setCellValue('Q'.$i, $row['Kecamatan']);
	$sheet->setCellValue('R'.$i, $row['Kode pos']);	
	$sheet->setCellValue('S'.$i, $row['Lintang']);	
	$sheet->setCellValue('T'.$i, $row['Bujur']);
	$sheet->setCellValue('U'.$i, $row['Tempat Tinggal']);
	$sheet->setCellValue('V'.$i, $row['Mode Transport']);	
	$sheet->setCellValue('W'.$i, $row['Nomor KKS']);
	$sheet->setCellValue('X'.$i, $row['Anak ke']);	
	$sheet->setCellValue('Y'.$i, $row['KPS/KPH']);	
	$sheet->setCellValue('Z'.$i, $row['No. KPS/KPH']);						
	$i++;
}

$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1;
$sheet->getStyle('A1:Z'.$i)->applyFromArray($styleArray);


$writer = new Xlsx($spreadsheet);
$writer->save('SiswaBaru.xlsx');
?>