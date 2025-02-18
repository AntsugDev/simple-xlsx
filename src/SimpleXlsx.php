<?php

namespace Anton\SimpleXlsx;

use Carbon\Carbon;
use Illuminate\Config\Repository as Config;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use ZipArchive;

class SimpleXlsx
{
    protected array $header;

    protected array $body;

    protected ?array $sheets;

    protected Spreadsheet $spreadsheet;

    protected string $path;

    protected string $title;

    protected ?int $defaultRowBody;
    protected ?int $defaultRowHeader;

    protected string|array|null $extradati;

    protected ?int $len;

    protected ?array $extracolor;

    protected Config $driver;

    /**
     * @param array $header
     * @param string $title
     * @param array|null $sheets
     * @throws \Exception
     */
    public function __construct(array $header, string $title, ?array $sheets,?int $defaultRowHeader =1,string|array|null $extradati = null,?string $pathBase = null,?int $len = null, ?array $extraColor = null )
    {
        $this->header = $header;
        $this->sheets = $sheets;

        $this->extracolor = $extraColor;

        if(!is_null($this->extracolor) && !array_key_exists("background",$this->extracolor) && !array_key_exists("color",$this->extracolor))
            throw new \Exception("If the last argument is evaluated, this is an array that must contain the fields background(background color) and color(font color)",400);

        if(!is_null($this->extracolor) && array_key_exists("align",$this->extracolor)  && !in_array($this->extracolor['align'],['center','left','right']))
            throw new \Exception("If the last argument is evaluated and it contains the align field, this can be 'center','left','right'",400);

        if(!is_null($this->sheets) && count($this->sheets) !== count($this->header))
            throw new \Exception("Inconsistency between nr.sheet and nr. of sheet present in the headers",400);

        $this->title  = $title;
        $this->defaultRowBody = ($defaultRowHeader+1);
        $this->defaultRowHeader = $defaultRowHeader;

        $this->extradati = $extradati;
        if(!is_null($this->extradati) && $this->defaultRowHeader === 1)
            throw new \Exception("If you want to show the extra data, the starting row of the table must be greater than 1",400);

        $this->setPath($pathBase);
        $this->spreadsheet = new Spreadsheet();
        $this->len      = $len;
    }
    const letterHeader = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

    public function setDriver(?string $pathSave = null): void
    {
        $pathSave = is_null($pathSave) ? str_replace('src','',__DIR__) . 'storage' : $pathSave;

        $this->driver = new Config(['filesystems' => [
            'default' => 'local',
            'disks' => [
                'local' => [
                    'driver' => 'local',
                    'root' => $pathSave,
                ],
            ],
        ]]);
    }



    public function extraDati($sheet, int $len): void
    {
        if($this->defaultRowHeader > 1 && !is_null($this->extradati)){
            $styleExtract = [
                'font' => [
                    'bold' => true,
                    'color' => ['rgb' => (!is_null($this->extracolor) ? $this->extracolor['color'] : 'FFFFFF')], // Red text
                    'size' => 13,
                    'name' => 'Calibri',
                ],
                'alignment' => [
                    'horizontal' => (!is_null($this->extracolor) && array_key_exists("align",$this->extracolor) ? $this->extracolor['align'] :   \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER),
                ],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['rgb' => (!is_null($this->extracolor) ? $this->extracolor['background'] :'FFC107' )],
                ],
                'borders' => [
                    'outline' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
                        'color' => ['rgb' => (!is_null($this->extracolor) ? $this->extracolor['color'] : 'FED91F')],
                    ],
                ]
            ];
            $rangeVerticalMax = self::letterHeader[( !is_null($this->len) ? ($this->len-1):  ($len-1))];
            for($i = 1; $i <= ($this->defaultRowHeader-1); $i ++){
                $sheet->mergeCells('A'.$i.':'.$rangeVerticalMax.$i);
            }
            if(is_string($this->extradati))
                $sheet->setCellValue('A1',$this->extradati);
            else{
                if(count($this->extradati) > 0)
                    foreach ($this->extradati as $key => $data){
                        $sheet->setCellValue('A'.($key+1),$data);
                    }
                else {
                    $this->defaultRowHeader = 1;
                    $this->defaultRowBody   = 2;
                }
            }
            $sheet->getStyle('A1:'.$rangeVerticalMax.($this->defaultRowHeader-1))->applyFromArray($styleExtract);
        }
    }

    /**
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function setSpreadsheet(?string $background = 'FFFFFF',?string $color = '000000',?string $name = null): int
    {
        $this->setSheet($name);
        $this->setHeader($background,$color);
        return $this->defaultRowBody;
    }

    public static function row (): ?int
    {
        return self::getDefaultRowBody();
    }

    public function getDefaultRowBody(): ?int
    {
        return $this->defaultRowBody;
    }

    public function setDefaultRowBody(?int $defaultRowBody): void
    {
        $this->defaultRowBody = $defaultRowBody;
    }

    public function getDefaultRowHeader(): ?int
    {
        return $this->defaultRowHeader;
    }

    public function setDefaultRowHeader(?int $defaultRowHeader): void
    {
        $this->defaultRowHeader = $defaultRowHeader;
    }

    public function getPath(): string
    {
        return $this->path;
    }

    /**
     * @throws Exception
     */
    protected function setSheet(?string $name = null): void
    {
        if(!is_null($this->sheets) && count($this->sheets) > 0){
            foreach ($this->sheets as $k => $sheet){
                $workSheet = new Worksheet();
                if(strlen($sheet) > 31)
                    $sheet = strtoupper(str_pad(substr($sheet,0,25),3,'.'));
                $workSheet->setTitle($sheet);
                $this->spreadsheet->addSheet($workSheet,$k);
            }
        }else{
            $workSheet = new Worksheet();
            $title = Carbon::now()->format('Y-m-d');
            if(!is_null($name))
                $title = $name;
            if(strlen($title) > 31)
                $title = strtoupper(str_pad(substr($title,0,25),3,'.'));
            $workSheet->setTitle($title);
            $this->spreadsheet->addSheet($workSheet,0);
        }
    }

    protected function styleHeader (?string $background ='FFFFFF',?string $color = '000000' ): array
    {
        return  [
            'font' => [
                'bold' => true,
                'color' => ['rgb' => $color], // Red text
                'size' => 13,
                'name' => 'Calibri',
                'align'=> 'center'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $background],
            ],
        ];
    }

    /**
     * @throws \Exception
     */
    public function styleBody (?bool $color= null, ?string $align = null, ?array $fill = null): array
    {
        if(!is_null($align)  && (strcmp($align,'center') !== 0 && strcmp($align,'left') !== 0 && strcmp($align,'right') !== 0))
            throw new \Exception("The align argument can be 'center','left','right'",400);

        if(!is_null($fill)){
            if(!array_key_exists('color',$fill) || !array_key_exists('bold',$fill))
                throw new \Exception("The fill field must be an array that must contain at least one of the non-null 'color', 'bold' keys.",400);

        }

        $fillColor = (is_null($color) ? 'FFFFFF' : ($color ? 'F2F3F4' : 'EAEDED'));
        if(!is_null($fill) && array_key_exists('color',$fill) && !is_null($fill['color'])){
            $fillColor = $fill['color'];
        }

        return [
            'font' => [
                'bold' => !is_null($fill) && array_key_exists('bold',$fill) && !is_null($fill['bold']) ? $fill['bold'] :false,
                'color' => ['rgb' => '000000'], // Red text
                'size' => 11,
                'name' => 'Calibri',
            ],
            'alignment' => [
                'horizontal' => !is_null($align) ? $align : Alignment::HORIZONTAL_LEFT,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $fillColor],
            ],
        ];
    }


    /**
     * @throws Exception
     * @throws \Exception
     */
    protected function setHeader(?string $background = 'FFFFFF',?string $color = '000000'): void
    {
        foreach ($this->header as $key => $h){
            $sheet = $this->spreadsheet->getSheet($key);
            $this->spreadsheet->getActiveSheet();
            if(!is_array($h))
                throw new \Exception("Array not correct for extra data",406);
            $this->extraDati($sheet, count($h));
            for($i = 0 ; $i < count($h); $i++) {
                $sheet->setCellValue(self::letterHeader[$i].$this->getDefaultRowHeader(),strtoupper($h[$i]));
                $sheet->getColumnDimension(self::letterHeader[$i])->setWidth(40);
            }
            $sheet->getStyle(self::letterHeader[0].$this->getDefaultRowHeader().':'.self::letterHeader[count($h)-1].$this->getDefaultRowHeader())->applyFromArray($this->styleHeader($background,$color));
            $sheet->setAutoFilter(self::letterHeader[0].$this->getDefaultRowHeader().':'.self::letterHeader[count($h)-1].$this->getDefaultRowHeader());
        }
    }


    /**
     * @throws Exception
     * @throws \Exception
     */
    public function setBodyCell(int $sheetKey, int $index, int $row, mixed $value, ?bool $color,?string $align= null, ?string $numberFormat = null, ?array $fill = null): void
    {

        $sheet = $this->spreadsheet->getSheet($sheetKey);
        $this->spreadsheet->getActiveSheet();
        $value = preg_replace(['/à/', '/è/', '/ì/', '/ò/', '/ù/'], ["A'", "E'", "I'", "O'", "U'"], $value);
        $sheet->setCellValue(self::letterHeader[$index].$row,$value)->getStyle(self::letterHeader[$index].$row)->applyFromArray($this->styleBody($color, $align,$fill))
            ->getBorders()->getBottom()
            ->setBorderStyle(Border::BORDER_THICK)
            ->setColor(new Color(Color::COLOR_WHITE));
        if(!is_null($numberFormat))
            $sheet->getStyle(self::letterHeader[$index].$row)->getNumberFormat()->setFormatCode($numberFormat);
        else
            $sheet->getStyle(self::letterHeader[$index].$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_GENERAL);
    }



    protected function setPath(string $path = null ): void
    {
        $title = strtoupper($this->title);
        $this->path = $path.'/'.$title.'.xlsx';
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws Exception
     */
    public function save(?ZipArchive $zipArchive = null):void
    {
        if($this->spreadsheet->getSheet(count($this->header))) {
            $this->spreadsheet->removeSheetByIndex(count($this->header));
        }
        $this->spreadsheet->setActiveSheetIndex(0);
        $xlsx = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
        $xlsx->save($this->path);
        if(!is_null($zipArchive)) {
            $zipArchive->addFile($this->path, basename($this->path));
        }
    }
	
	  public function setMergeCells(int $index, string $from, string $to,string $dati, ?array $style = null): void
    {
        $sheet = $this->spreadsheet->getSheet($index);
        $sheet->mergeCells($from.':'.$to);
        $sheet->setCellValue($from,$dati);
        if(!is_null($style)) $sheet->getStyle($from)->applyFromArray($style);
    }

    /**
     * @throws Exception
     */
    public function getDimension(int $index, string $column, ?int $width = 40 ): void
    {
        $sheet = $this->spreadsheet->getSheet($index);
        $sheet->getColumnDimension($column)->setWidth($width);
    }

}